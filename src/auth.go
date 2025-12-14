package main

import (
	"context"
	"encoding/json"
	"fmt"
	"net/http"
	"os"
	"path/filepath"

	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/drive/v3"
	"google.golang.org/api/option"
	"google.golang.org/api/slides/v1"
)

const (
	credentialsFile = "google_credentials.json"
	tokenFile       = "google_token.json"
)

var scopes = []string{
	slides.PresentationsScope,
	drive.DriveFileScope,
	"https://www.googleapis.com/auth/cloud-translation", // Translation API
}

// getCredentialsPath returns the path to credentials directory
func getCredentialsPath() (string, error) {
	home, err := os.UserHomeDir()
	if err != nil {
		return "", fmt.Errorf("unable to get home directory: %w", err)
	}
	return filepath.Join(home, ".credentials"), nil
}

// getClient retrieves an OAuth2 HTTP client
func getClient(ctx context.Context) (*http.Client, error) {
	credentialsPath, err := getCredentialsPath()
	if err != nil {
		return nil, err
	}

	credPath := filepath.Join(credentialsPath, credentialsFile)
	tokenPath := filepath.Join(credentialsPath, tokenFile)

	// Read credentials file
	b, err := os.ReadFile(credPath)
	if err != nil {
		return nil, fmt.Errorf("unable to read credentials file %s: %w\n"+
			"See README.md for setup instructions", credPath, err)
	}

	// Parse credentials
	config, err := google.ConfigFromJSON(b, scopes...)
	if err != nil {
		return nil, fmt.Errorf("unable to parse credentials: %w", err)
	}

	// Try to load token from file
	token, err := tokenFromFile(tokenPath)
	if err != nil {
		// Get new token from user
		token, err = getTokenFromWeb(config)
		if err != nil {
			return nil, err
		}
		// Save token
		if err := saveToken(tokenPath, token); err != nil {
			return nil, fmt.Errorf("unable to save token: %w", err)
		}
	}

	return config.Client(ctx, token), nil
}

// getSlidesService creates an authenticated Slides service
func getSlidesService(ctx context.Context) (*slides.Service, error) {
	client, err := getClient(ctx)
	if err != nil {
		return nil, err
	}

	service, err := slides.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return nil, fmt.Errorf("unable to create Slides service: %w", err)
	}

	return service, nil
}

// getDriveService creates an authenticated Drive service
func getDriveService(ctx context.Context) (*drive.Service, error) {
	client, err := getClient(ctx)
	if err != nil {
		return nil, err
	}

	service, err := drive.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return nil, fmt.Errorf("unable to create Drive service: %w", err)
	}

	return service, nil
}

// getTokenFromWeb requests a token from the web, then returns the retrieved token
func getTokenFromWeb(config *oauth2.Config) (*oauth2.Token, error) {
	authURL := config.AuthCodeURL("state-token", oauth2.AccessTypeOffline)
	fmt.Printf("Go to the following link in your browser:\n%v\n\n", authURL)
	fmt.Printf("Enter authorization code: ")

	var authCode string
	if _, err := fmt.Scan(&authCode); err != nil {
		return nil, fmt.Errorf("unable to read authorization code: %w", err)
	}

	token, err := config.Exchange(context.TODO(), authCode)
	if err != nil {
		return nil, fmt.Errorf("unable to retrieve token from web: %w", err)
	}

	return token, nil
}

// tokenFromFile retrieves a token from a local file
func tokenFromFile(file string) (*oauth2.Token, error) {
	f, err := os.Open(file)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	token := &oauth2.Token{}
	err = json.NewDecoder(f).Decode(token)
	return token, err
}

// saveToken saves a token to a file path
func saveToken(path string, token *oauth2.Token) error {
	fmt.Fprintf(os.Stderr, "Saving credentials to: %s\n", path)

	// Ensure directory exists
	if err := os.MkdirAll(filepath.Dir(path), 0700); err != nil {
		return err
	}

	f, err := os.OpenFile(path, os.O_RDWR|os.O_CREATE|os.O_TRUNC, 0600)
	if err != nil {
		return err
	}
	defer f.Close()

	return json.NewEncoder(f).Encode(token)
}
