set -e

echo "===== Node & npm ====="
node -v
npm -v

echo "===== Preparing input_files ====="
mkdir -p input_files/global_users
mkdir -p input_files/project_users
mkdir -p input_files/repo_users

rm -f input_files/global_users/*
rm -f input_files/project_users/*
rm -f input_files/repo_users/*

echo "===== Copying uploaded CSV files ====="

echo "GLOBAL_USERS_FILE = $GLOBAL_USERS_FILE"
echo "PROJECT_USERS_FILE = $PROJECT_USERS_FILE"
echo "REPO_USERS_FILE = $REPO_USERS_FILE"

# Global users
if [ -f "$GLOBAL_USERS_FILE" ]; then
  cp "$GLOBAL_USERS_FILE" input_files/global_users/
else
  echo "ERROR: GLOBAL_USERS_FILE not provided"
  exit 1
fi

# Project users
if [ -f "$PROJECT_USERS_FILE" ]; then
  cp "$PROJECT_USERS_FILE" input_files/project_users/
else
  echo "ERROR: PROJECT_USERS_FILE not provided"
  exit 1
fi

# Repo users
if [ -f "$REPO_USERS_FILE" ]; then
  cp "$REPO_USERS_FILE" input_files/repo_users/
else
  echo "ERROR: REPO_USERS_FILE not provided"
  exit 1
fi

echo "===== input_files structure ====="
find input_files -type f

echo "===== Installing deps ====="
npm install

echo "===== Running app ====="
node src/index.js