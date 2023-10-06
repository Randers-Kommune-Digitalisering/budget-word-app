#!/bin/sh

tag_lower=$(echo $1 | tr '[:upper:]' '[:lower:]')
# Build docker images.
docker build --tag $tag_lower --label org.opencontainers.image.source=$2 ./