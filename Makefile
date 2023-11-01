BINARY_NAME := xlsxsql
SRCS := $(shell git ls-files '*.go')
LDFLAGS := "-X main.version=$(shell git describe --tags --abbrev=0 --always) -X main.revision=$(shell git rev-parse --short HEAD)"

all: build

test: $(SRCS)
	go test ./...

build: $(BINARY_NAME)

$(BINARY_NAME): $(SRCS)
	go build -ldflags $(LDFLAGS) -o $(BINARY_NAME) ./cmd/xlsxsql

install:
	go install -ldflags $(LDFLAGS) ./cmd/xlsxsql

clean:
	rm -f $(BINARY_NAME)

.PHONY: all test build install clean