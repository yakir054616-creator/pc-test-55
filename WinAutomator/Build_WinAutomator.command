#!/bin/bash
cd "$(dirname "$0")"

echo "Building WinAutomator Modern UI (win-x64)..."
dotnet publish "WinAutomator.csproj" -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o ./Publish

if [ $? -eq 0 ]; then
    echo "========================================="
    echo "SUCCESS! The .exe is ready in the 'Publish' folder."
    echo "========================================="
else
    echo "========================================="
    echo "BUILD FAILED. Please check the errors above."
    echo "========================================="
fi

read -p "Press any key to close this window..."
