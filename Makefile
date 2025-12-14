.PHONY: help install test test-converter test-vbs test-vbs-docker convert clean docker-build docker-test docker-clean

# Default target
help:
	@echo "VBA-CI-Sample Makefile"
	@echo ""
	@echo "Usage:"
	@echo "  make install        - Install Node.js dependencies"
	@echo "  make test           - Run all tests (converter only on non-Windows)"
	@echo "  make test-converter - Run Node.js converter tests"
	@echo "  make test-vbs       - Run VBS tests (Windows only)"
	@echo "  make convert        - Convert VBA to VBS"
	@echo "  make clean          - Clean generated files"
	@echo ""
	@echo "Docker commands:"
	@echo "  make docker-build   - Build Docker image for VBS testing"
	@echo "  make docker-test    - Run tests in Docker container"
	@echo "  make docker-clean   - Remove Docker containers and images"
	@echo ""

# Install dependencies
install:
	cd build && npm ci

# Run converter tests (Node.js - cross-platform)
test-converter:
	cd build && npm test

# Convert VBA to VBS
convert:
	node build/converter/index.js --input ./src ./test --output ./build/vbs/generated

# Run VBS tests (Windows only - uses cscript)
test-vbs:
ifeq ($(OS),Windows_NT)
	cscript //nologo build/vbs/run-tests.vbs
else
	@echo "VBS tests require Windows. Use 'make docker-test' for Wine-based testing."
	@exit 1
endif

# Run VBS tests inside Docker (using Wine)
test-vbs-docker:
	cd build && npm ci
	node build/converter/index.js --input ./src ./test --output ./build/vbs/generated
	wine cscript //nologo build/vbs/run-tests.vbs

# Run all tests
test: test-converter
ifeq ($(OS),Windows_NT)
	$(MAKE) convert
	$(MAKE) test-vbs
endif

# Full test in Docker
test-all: test-converter convert test-vbs-docker

# Clean generated files
clean:
	rm -rf build/vbs/generated/*
	rm -f build/vbs/vba-compat.vbs

# Docker commands
docker-build:
	docker compose build vbs-test

docker-test:
	docker compose run --rm vbs-test make test-all

docker-converter-test:
	docker compose run --rm converter-test

docker-clean:
	docker compose down --rmi local --volumes --remove-orphans
	docker image prune -f
