#!/bin/bash
# TB-GL Linker Deployment Script
# Usage: ./deploy.sh [command]
#
# Commands:
#   build    - Build the Docker image
#   start    - Start the container (production)
#   stop     - Stop the container
#   restart  - Restart the container
#   logs     - View container logs
#   dev      - Start in development mode (hot reload)
#   update   - Pull latest changes and restart
#   status   - Check container status

set -e

CONTAINER_NAME="tb-gl-linker"
IMAGE_NAME="tb-gl-linker:latest"

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

print_status() {
    echo -e "${GREEN}[INFO]${NC} $1"
}

print_warning() {
    echo -e "${YELLOW}[WARN]${NC} $1"
}

print_error() {
    echo -e "${RED}[ERROR]${NC} $1"
}

case "${1:-start}" in
    build)
        print_status "Building Docker image..."
        docker-compose build
        print_status "Build complete!"
        ;;

    start)
        print_status "Starting TB-GL Linker..."
        docker-compose up -d tb-gl-linker
        print_status "Container started!"
        print_status "Access the app at: http://localhost:8530"
        ;;

    stop)
        print_status "Stopping TB-GL Linker..."
        docker-compose down
        print_status "Container stopped!"
        ;;

    restart)
        print_status "Restarting TB-GL Linker..."
        docker-compose restart tb-gl-linker
        print_status "Container restarted!"
        ;;

    logs)
        print_status "Showing logs (Ctrl+C to exit)..."
        docker-compose logs -f tb-gl-linker
        ;;

    dev)
        print_status "Starting in development mode..."
        docker-compose --profile dev up -d tb-gl-linker-dev
        print_status "Development server started!"
        print_status "Access the app at: http://localhost:8531"
        print_status "Changes to source files will auto-reload"
        ;;

    update)
        print_status "Pulling latest changes..."
        git pull
        print_status "Rebuilding image..."
        docker-compose build
        print_status "Restarting container..."
        docker-compose up -d tb-gl-linker
        print_status "Update complete!"
        print_status "Access the app at: http://localhost:8530"
        ;;

    status)
        print_status "Container status:"
        docker-compose ps
        echo ""
        print_status "Health check:"
        if curl -s -f http://localhost:8530/_stcore/health > /dev/null 2>&1; then
            echo -e "${GREEN}✓ App is healthy and running${NC}"
        else
            echo -e "${RED}✗ App is not responding${NC}"
        fi
        ;;

    *)
        echo "TB-GL Linker Deployment Script"
        echo ""
        echo "Usage: ./deploy.sh [command]"
        echo ""
        echo "Commands:"
        echo "  build    - Build the Docker image"
        echo "  start    - Start the container (production)"
        echo "  stop     - Stop the container"
        echo "  restart  - Restart the container"
        echo "  logs     - View container logs"
        echo "  dev      - Start in development mode (hot reload)"
        echo "  update   - Pull latest changes and restart"
        echo "  status   - Check container status"
        echo ""
        echo "Examples:"
        echo "  ./deploy.sh build   # Build the image"
        echo "  ./deploy.sh start   # Start production server"
        echo "  ./deploy.sh update  # Pull changes and redeploy"
        ;;
esac
