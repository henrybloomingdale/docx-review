# docx-review Makefile
# Build a single native binary using .NET 8 + Open XML SDK
#
# Usage:
#   make              # Build for current platform
#   make install      # Build + install to /usr/local/bin
#   make all          # Build for macOS ARM64, macOS x64, and Linux x64
#   make docker       # Build Docker image
#   make test         # Run test against sample document
#   make clean        # Remove build artifacts

BINARY_NAME  := docx-review
VERSION      := 1.3.5
BUILD_DIR    := build
INSTALL_DIR  := /usr/local/bin
GH_REPO      := drpedapati/docx-review
HOMEBREW_TAPS := drpedapati/homebrew-tools drpedapati/homebrew-tap

# Detect current platform
UNAME_S := $(shell uname -s)
UNAME_M := $(shell uname -m)

ifeq ($(UNAME_S),Darwin)
  ifeq ($(UNAME_M),arm64)
    CURRENT_RID := osx-arm64
  else
    CURRENT_RID := osx-x64
  endif
else
  CURRENT_RID := linux-x64
endif

# .NET publish flags
PUBLISH_FLAGS := -c Release \
  --self-contained \
  -p:PublishSingleFile=true \
  -p:EnableCompressionInSingleFile=true \
  -p:IncludeNativeLibrariesForSelfExtract=true \
  -p:PublishTrimmed=true \
  -p:TrimMode=link \
  -p:SuppressTrimAnalysisWarnings=true

# Release asset names (must match homebrew formula URLs)
RELEASE_ASSETS := \
  $(BUILD_DIR)/$(BINARY_NAME)-darwin-arm64 \
  $(BUILD_DIR)/$(BINARY_NAME)-darwin-amd64 \
  $(BUILD_DIR)/$(BINARY_NAME)-linux-amd64 \
  $(BUILD_DIR)/$(BINARY_NAME)-linux-arm64

.PHONY: build install all docker test test-create test-comment-update clean help release release-assets update-taps

## build: Build single-file binary for current platform
build:
	@echo "Building $(BINARY_NAME) for $(CURRENT_RID)..."
	@mkdir -p $(BUILD_DIR)
	dotnet publish $(PUBLISH_FLAGS) -r $(CURRENT_RID) -o $(BUILD_DIR)/$(CURRENT_RID)
	@cp $(BUILD_DIR)/$(CURRENT_RID)/$(BINARY_NAME) $(BUILD_DIR)/$(BINARY_NAME)
	@echo ""
	@echo "✅ Built: $(BUILD_DIR)/$(BINARY_NAME)"
	@ls -lh $(BUILD_DIR)/$(BINARY_NAME)

## install: Build and install to /usr/local/bin
install: build
	@echo "Installing to $(INSTALL_DIR)/$(BINARY_NAME)..."
	@cp $(BUILD_DIR)/$(BINARY_NAME) $(INSTALL_DIR)/$(BINARY_NAME)
	@chmod +x $(INSTALL_DIR)/$(BINARY_NAME)
	@echo "✅ Installed: $(INSTALL_DIR)/$(BINARY_NAME)"

## uninstall: Remove from /usr/local/bin
uninstall:
	@rm -f $(INSTALL_DIR)/$(BINARY_NAME)
	@echo "Removed $(INSTALL_DIR)/$(BINARY_NAME)"

## all: Build for all platforms (macOS ARM64, macOS x64, Linux x64)
all:
	@echo "Building for all platforms..."
	@mkdir -p $(BUILD_DIR)
	@for rid in osx-arm64 osx-x64 linux-x64 linux-arm64; do \
		echo ""; \
		echo "→ Building for $$rid..."; \
		dotnet publish $(PUBLISH_FLAGS) -r $$rid -o $(BUILD_DIR)/$$rid; \
		echo "  ✅ $(BUILD_DIR)/$$rid/$(BINARY_NAME)"; \
	done
	@echo ""
	@echo "All builds complete:"
	@ls -lh $(BUILD_DIR)/osx-arm64/$(BINARY_NAME) $(BUILD_DIR)/osx-x64/$(BINARY_NAME) $(BUILD_DIR)/linux-x64/$(BINARY_NAME) $(BUILD_DIR)/linux-arm64/$(BINARY_NAME) 2>/dev/null

## docker: Build Docker image
docker:
	docker build -t $(BINARY_NAME) .

## test: Run test against the example manifest
test: build
	@echo "Running test..."
	@if [ ! -f examples/sample-edits.json ]; then \
		echo "Error: examples/sample-edits.json not found"; \
		exit 1; \
	fi
	@if [ ! -f "$(TEST_DOC)" ]; then \
		echo "Usage: make test TEST_DOC=/path/to/document.docx"; \
		echo "  e.g. make test TEST_DOC=~/Dropbox/Henry\\ Projects/mbhi-ai-proposals/docs/Cognitive_Choreography_v7_npjDigMed.docx"; \
		exit 1; \
	fi
	$(BUILD_DIR)/$(BINARY_NAME) "$(TEST_DOC)" examples/sample-edits.json -o $(BUILD_DIR)/test_output.docx
	@echo ""
	@ls -lh $(BUILD_DIR)/test_output.docx

## test-dry: Dry-run test (no modifications)
test-dry: build
	@if [ -f "$(TEST_DOC)" ]; then \
		$(BUILD_DIR)/$(BINARY_NAME) "$(TEST_DOC)" examples/sample-edits.json --dry-run; \
	else \
		echo "Usage: make test-dry TEST_DOC=/path/to/document.docx"; \
	fi

## test-create: Test create mode
test-create: build
	@echo "Testing --create mode..."
	$(BUILD_DIR)/$(BINARY_NAME) --create -o $(BUILD_DIR)/test_created.docx --json
	@echo ""
	@echo "Testing --create dry-run..."
	$(BUILD_DIR)/$(BINARY_NAME) --create --dry-run --json
	@echo ""
	@ls -lh $(BUILD_DIR)/test_created.docx
	@rm -f $(BUILD_DIR)/test_created.docx
	@echo "✅ Create tests passed"

## test-comment-update: Integration test for comment update op by ID
test-comment-update:
	@bash tests/test-comment-update.sh

## release: Build all platforms, create GitHub release, update Homebrew taps
release: all release-assets
	@echo ""
	@echo "Creating GitHub release v$(VERSION)..."
	gh release create v$(VERSION) $(RELEASE_ASSETS) \
		--title "v$(VERSION)" \
		--notes "Release v$(VERSION)" \
		--repo $(GH_REPO)
	@echo "✅ GitHub release v$(VERSION) created"
	@echo ""
	@$(MAKE) update-taps

## release-assets: Copy platform binaries with release naming convention
release-assets:
	@cp $(BUILD_DIR)/osx-arm64/$(BINARY_NAME)   $(BUILD_DIR)/$(BINARY_NAME)-darwin-arm64
	@cp $(BUILD_DIR)/osx-x64/$(BINARY_NAME)     $(BUILD_DIR)/$(BINARY_NAME)-darwin-amd64
	@cp $(BUILD_DIR)/linux-x64/$(BINARY_NAME)   $(BUILD_DIR)/$(BINARY_NAME)-linux-amd64
	@cp $(BUILD_DIR)/linux-arm64/$(BINARY_NAME)  $(BUILD_DIR)/$(BINARY_NAME)-linux-arm64
	@echo "Release assets:"
	@ls -lh $(RELEASE_ASSETS)

## update-taps: Update all Homebrew tap formulas with current version and SHA256s
update-taps: release-assets
	@ARM64_SHA=$$(shasum -a 256 $(BUILD_DIR)/$(BINARY_NAME)-darwin-arm64 | cut -d' ' -f1); \
	AMD64_SHA=$$(shasum -a 256 $(BUILD_DIR)/$(BINARY_NAME)-darwin-amd64 | cut -d' ' -f1); \
	LINUX_AMD64_SHA=$$(shasum -a 256 $(BUILD_DIR)/$(BINARY_NAME)-linux-amd64 | cut -d' ' -f1); \
	LINUX_ARM64_SHA=$$(shasum -a 256 $(BUILD_DIR)/$(BINARY_NAME)-linux-arm64 | cut -d' ' -f1); \
	echo "SHA256 hashes:"; \
	echo "  darwin-arm64:  $$ARM64_SHA"; \
	echo "  darwin-amd64:  $$AMD64_SHA"; \
	echo "  linux-amd64:   $$LINUX_AMD64_SHA"; \
	echo "  linux-arm64:   $$LINUX_ARM64_SHA"; \
	echo ""; \
	for tap in $(HOMEBREW_TAPS); do \
		echo "→ Updating $$tap..."; \
		TMPDIR=$$(mktemp -d); \
		gh repo clone $$tap $$TMPDIR 2>/dev/null; \
		FORMULA=$$TMPDIR/Formula/$(BINARY_NAME).rb; \
		if [ ! -f "$$FORMULA" ]; then \
			echo "  ⚠️  No formula found in $$tap, skipping"; \
			rm -rf $$TMPDIR; \
			continue; \
		fi; \
		sed -i '' "s/version \"[^\"]*\"/version \"$(VERSION)\"/" $$FORMULA; \
		sed -i '' "s|download/v[^/]*/|download/v$(VERSION)/|g" $$FORMULA; \
		sed -i '' "s|docx-review [0-9][0-9]*\.[0-9][0-9]*\.[0-9][0-9]*|docx-review $(VERSION)|" $$FORMULA; \
		darwin_arm64_old=$$(grep -A1 'arm64\|on_arm' $$FORMULA | grep sha256 | head -1 | sed 's/.*"\(.*\)".*/\1/'); \
		darwin_amd64_old=$$(grep -A1 'intel\|on_intel\|CPU.arm' $$FORMULA | grep sha256 | head -1 | sed 's/.*"\(.*\)".*/\1/'); \
		linux_arm64_old=$$(grep -B0 -A3 'on_linux' $$FORMULA | grep -A1 'arm64\|on_arm' | grep sha256 | head -1 | sed 's/.*"\(.*\)".*/\1/'); \
		linux_amd64_old=$$(grep -B0 -A3 'on_linux' $$FORMULA | grep -A1 'intel\|on_intel' | grep sha256 | head -1 | sed 's/.*"\(.*\)".*/\1/'); \
		if [ -n "$$darwin_arm64_old" ]; then sed -i '' "s/$$darwin_arm64_old/$$ARM64_SHA/" $$FORMULA; fi; \
		if [ -n "$$darwin_amd64_old" ]; then sed -i '' "s/$$darwin_amd64_old/$$AMD64_SHA/" $$FORMULA; fi; \
		if [ -n "$$linux_arm64_old" ]; then sed -i '' "s/$$linux_arm64_old/$$LINUX_ARM64_SHA/" $$FORMULA; fi; \
		if [ -n "$$linux_amd64_old" ]; then sed -i '' "s/$$linux_amd64_old/$$LINUX_AMD64_SHA/" $$FORMULA; fi; \
		cd $$TMPDIR && git add -A && git commit -m "Update $(BINARY_NAME) to v$(VERSION)" && git push origin main; \
		cd - > /dev/null; \
		rm -rf $$TMPDIR; \
		echo "  ✅ $$tap updated to v$(VERSION)"; \
	done
	@echo ""
	@echo "✅ All Homebrew taps updated. Run: brew update && brew upgrade $(BINARY_NAME)"

## clean: Remove build artifacts
clean:
	@rm -rf $(BUILD_DIR) bin obj
	@echo "Cleaned build artifacts"

## help: Show this help
help:
	@echo "docx-review $(VERSION) — Word document review tool"
	@echo ""
	@echo "Targets:"
	@grep -E '^## ' Makefile | sed 's/## /  /' | column -t -s ':'
	@echo ""
	@echo "Examples:"
	@echo "  make                          # Build for $(CURRENT_RID)"
	@echo "  make install                  # Build + install to $(INSTALL_DIR)"
	@echo "  make all                      # Cross-compile all platforms"
	@echo "  make test TEST_DOC=paper.docx # Run test"
	@echo "  make clean                    # Remove artifacts"
