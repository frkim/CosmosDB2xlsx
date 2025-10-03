#!/bin/bash
# Script to publish CosmosDB2xlsx for multiple platforms

echo "Publishing CosmosDB2xlsx..."

# Publish for Windows (x64)
echo "Building for Windows x64..."
dotnet publish CosmosDB2xlsx/CosmosDB2xlsx.csproj -c Release -r win-x64 --self-contained false -o ./publish/win-x64

# Publish for Linux (x64)
echo "Building for Linux x64..."
dotnet publish CosmosDB2xlsx/CosmosDB2xlsx.csproj -c Release -r linux-x64 --self-contained false -o ./publish/linux-x64

# Publish for macOS (x64)
echo "Building for macOS x64..."
dotnet publish CosmosDB2xlsx/CosmosDB2xlsx.csproj -c Release -r osx-x64 --self-contained false -o ./publish/osx-x64

# Publish for macOS (ARM64)
echo "Building for macOS ARM64..."
dotnet publish CosmosDB2xlsx/CosmosDB2xlsx.csproj -c Release -r osx-arm64 --self-contained false -o ./publish/osx-arm64

echo "Publishing complete! Binaries are in the ./publish directory"
