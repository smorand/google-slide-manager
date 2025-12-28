package auth

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
	credentialsFileName     = "credentials.json"
	tokenFileName           = "token.json"
	translationAPIScope     = "https://www.googleapis.com/auth/cloud-translation"
)

var scopes = []string{
	slides.PresentationsScope,
	drive.DriveScope,
	translationAPIScope,
}

// GetCredentialsPath returns the path to the credentials directory.
func GetCredentialsPath() (string, error) {
	homeDir, err := os.UserHomeDir()
	if err != nil {
		return "", fmt.Errorf("unable to get home directory: %w", err)
	}
	return filepath.Join(homeDir, ".gdrive"), nil
}

// GetClient retrieves an OAuth2 HTTP client.
func GetClient(ctx context.Context) (*http.Client, error) {
	credentialsPath, err := GetCredentialsPath()
	if err != nil {
		return nil, err
	}

	credPath := filepath.Join(credentialsPath, credentialsFileName)
	tokenPath := filepath.Join(credentialsPath, tokenFileName)

	credentialsData, err := os.ReadFile(credPath)
	if err != nil {
		return nil, fmt.Errorf("unable to read credentials file %s: %w\nSee README.md for setup instructions", credPath, err)
	}

	config, err := google.ConfigFromJSON(credentialsData, scopes...)
	if err != nil {
		return nil, fmt.Errorf("unable to parse credentials: %w", err)
	}

	token, err := tokenFromFile(tokenPath)
	if err != nil {
		token, err = getTokenFromWeb(config)
		if err != nil {
			return nil, err
		}
		if err := saveToken(tokenPath, token); err != nil {
			return nil, fmt.Errorf("unable to save token: %w", err)
		}
	}

	return config.Client(ctx, token), nil
}

// GetDriveService creates an authenticated Drive service.
func GetDriveService(ctx context.Context) (*drive.Service, error) {
	client, err := GetClient(ctx)
	if err != nil {
		return nil, err
	}

	service, err := drive.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return nil, fmt.Errorf("unable to create Drive service: %w", err)
	}

	return service, nil
}

// GetSlidesService creates an authenticated Slides service.
func GetSlidesService(ctx context.Context) (*slides.Service, error) {
	client, err := GetClient(ctx)
	if err != nil {
		return nil, err
	}

	service, err := slides.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return nil, fmt.Errorf("unable to create Slides service: %w", err)
	}

	return service, nil
}

// getTokenFromWeb requests a token from the web through user authorization.
func getTokenFromWeb(config *oauth2.Config) (*oauth2.Token, error) {
	authURL := config.AuthCodeURL("state-token", oauth2.AccessTypeOffline)
	fmt.Printf("Go to the following link in your browser:\n%v\n\n", authURL)
	fmt.Printf("Enter authorization code: ")

	var authCode string
	if _, err := fmt.Scan(&authCode); err != nil {
		return nil, fmt.Errorf("unable to read authorization code: %w", err)
	}

	token, err := config.Exchange(context.Background(), authCode)
	if err != nil {
		return nil, fmt.Errorf("unable to retrieve token from web: %w", err)
	}

	return token, nil
}

// saveToken saves an OAuth2 token to a file path.
func saveToken(path string, token *oauth2.Token) error {
	fmt.Fprintf(os.Stderr, "Saving credentials to: %s\n", path)

	if err := os.MkdirAll(filepath.Dir(path), 0700); err != nil {
		return fmt.Errorf("unable to create credentials directory: %w", err)
	}

	file, err := os.OpenFile(path, os.O_RDWR|os.O_CREATE|os.O_TRUNC, 0600)
	if err != nil {
		return fmt.Errorf("unable to create token file: %w", err)
	}
	defer file.Close()

	if err := json.NewEncoder(file).Encode(token); err != nil {
		return fmt.Errorf("unable to encode token: %w", err)
	}

	return nil
}

// tokenFromFile retrieves an OAuth2 token from a local file.
func tokenFromFile(filePath string) (*oauth2.Token, error) {
	file, err := os.Open(filePath)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	token := &oauth2.Token{}
	if err := json.NewDecoder(file).Decode(token); err != nil {
		return nil, fmt.Errorf("unable to decode token: %w", err)
	}

	return token, nil
}
