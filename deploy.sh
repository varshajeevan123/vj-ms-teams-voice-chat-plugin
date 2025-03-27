#!/bin/bash

# Colors for output
GREEN='\033[0;32m'
RED='\033[0;31m'
NC='\033[0m'

echo -e "${GREEN}🚀 Starting deployment process...${NC}"

# Check if Node.js is installed
if ! command -v node &> /dev/null; then
    echo -e "${RED}❌ Node.js is not installed. Please install Node.js first.${NC}"
    exit 1
fi

# Install dependencies
echo -e "${GREEN}📦 Installing dependencies...${NC}"
npm install

# Create SSL directory if it doesn't exist
mkdir -p ssl

# Check if SSL certificates exist
if [ ! -f ssl/private.key ] || [ ! -f ssl/certificate.crt ]; then
    echo -e "${RED}❌ SSL certificates not found. Please generate them first.${NC}"
    echo -e "${GREEN}📝 Instructions for generating SSL certificates:${NC}"
    echo "1. Install OpenSSL"
    echo "2. Run: openssl req -x509 -nodes -days 365 -newkey rsa:2048 -keyout ssl/private.key -out ssl/certificate.crt"
    exit 1
fi

# Create .env file if it doesn't exist
if [ ! -f .env ]; then
    echo -e "${GREEN}📝 Creating .env file...${NC}"
    cat > .env << EOL
CLIENT_ID=your_client_id
CLIENT_SECRET=your_client_secret
TENANT_ID=your_tenant_id
SESSION_SECRET=your_session_secret
EOL
    echo -e "${RED}⚠️ Please update the .env file with your actual credentials${NC}"
    exit 1
fi

# Build the Teams app package
echo -e "${GREEN}🏗️ Building Teams app package...${NC}"
npm run build

# Start the server
echo -e "${GREEN}🚀 Starting the server...${NC}"
npm start

echo -e "${GREEN}✅ Deployment complete!${NC}"
echo -e "${GREEN}📝 Next steps:${NC}"
echo "1. Update the manifest.json with your server URL"
echo "2. Package the app using Teams Toolkit"
echo "3. Upload the app to Teams" 