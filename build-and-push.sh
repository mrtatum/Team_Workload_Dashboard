#!/usr/bin/env bash
# Build a multi-arch (amd64 + arm64) image and push it to GitHub Container Registry.
#
# Prereqs (one time):
#   1. Docker Desktop (or Docker Engine with buildx) installed and running.
#   2. A GitHub Personal Access Token with the `write:packages` scope.
#      https://github.com/settings/tokens  ->  generate (classic)  ->  write:packages
#   3. Log in once:
#        echo "$GH_TOKEN" | docker login ghcr.io -u <your-github-username> --password-stdin
#
# Usage:
#   GH_USER=your-github-username ./build-and-push.sh           # tags :latest
#   GH_USER=your-github-username TAG=v1.0.0 ./build-and-push.sh
set -euo pipefail

GH_USER="${GH_USER:?Set GH_USER to your GitHub username, e.g. GH_USER=tum ./build-and-push.sh}"
IMAGE="${IMAGE:-team-workload-dashboard}"
TAG="${TAG:-latest}"
PLATFORMS="${PLATFORMS:-linux/amd64,linux/arm64}"
REF="ghcr.io/${GH_USER}/${IMAGE}:${TAG}"

# Create (once) and select a buildx builder that supports multi-arch.
docker buildx inspect multiarch >/dev/null 2>&1 || docker buildx create --name multiarch --use
docker buildx use multiarch

echo "Building ${REF} for ${PLATFORMS} and pushing to GHCR..."
docker buildx build \
  --platform "${PLATFORMS}" \
  --tag "${REF}" \
  --push \
  .

echo "Done. Pulled with:  docker pull ${REF}"
echo "Run with:           docker run --env-file .env -p 8501:8501 ${REF}"
