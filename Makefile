.PHONY: help install test test-converter test-vbs convert clean docker-build docker-test docker-clean

# Default target
help:
	@echo "VBA-CI-Sample Makefile"
	@echo ""
	@echo "Usage:"
	@echo "  make install        - Install Node.js dependencies"
	@echo "  make test           - Run converter tests + convert (VBS tests on Windows only)"
	@echo "  make test-converter - Run Node.js converter tests"
	@echo "  make test-vbs       - Run VBS tests (Windows only)"
	@echo "  make convert        - Convert VBA to VBS"
	@echo "  make clean          - Clean generated files"
	@echo ""
	@echo "Docker commands (for cross-platform converter testing):"
	@echo "  make docker-build   - Build Docker image"
	@echo "  make docker-test    - Run converter tests in Docker"
	@echo "  make docker-clean   - Remove Docker containers and images"
	@echo ""
	@echo "Note: VBS tests require real Windows (GitHub Actions or local Windows)."
	@echo "      Wine on ARM64 macOS/Linux cannot run cscript reliably."
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
	@echo "ERROR: VBS tests require Windows."
	@echo "Use GitHub Actions or run on a Windows machine."
	@exit 1
endif

# Run all tests (converter + convert, VBS only on Windows)
test: test-converter convert
ifeq ($(OS),Windows_NT)
	$(MAKE) test-vbs
else
	@echo ""
	@echo "Converter tests passed. VBS tests skipped (requires Windows)."
	@echo "Push to GitHub to run full CI with VBS tests."
endif

# Clean generated files
clean:
	rm -rf build/vbs/generated/*
	rm -f build/vbs/vba-compat.vbs

# Docker commands (for converter tests only - VBS needs real Windows)
docker-build:
	docker compose build converter-test

docker-test:
	docker compose run --rm converter-test

docker-clean:
	docker compose down --rmi local --volumes --remove-orphans
	docker image prune -f
