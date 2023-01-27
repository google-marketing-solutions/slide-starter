#!/bin/bash

audit_type=$1
clasp_upload=$2
clasp_script_url=$3

# Check if the audit_type is valid
if [[ "$audit_type" != "app" && "$audit_type" != "web" && "$audit_type" != "ux" ]]; then
  echo "Error: Invalid audit type"
  exit 1
fi

# Function to check if logged into clasp
function check_login () {
  local login_status=$(clasp login --status)
  if [ "$login_status" == "You are not logged in." ]; then
    echo "Please log in using clasp login."
    exit 1
  fi
}

# Function to copy files from the main codebase into the clasp folder
function copy_files () {
  cp "../sls_core.js" "sls_core.js"
  cp "../sls_$audit_type.js" "sls_$audit_type.js"

  echo "Files successfully copied to clasp_uploads folder"
}

# Create the clasp_uploads folder if it doesn't already exist
if [ ! -d "clasp_uploads" ]; then
  mkdir clasp_uploads
fi

# Check if the user wants to upload to clasp
if [[ "$clasp_upload" == "--clasp-upload" ]]; then
  check_login
  if [ -z "$clasp_script_url" ]; then
    echo "Error: Script ID or URL not present"
    exit 1
  fi
  cd clasp_uploads
  clasp clone "$clasp_script_url"
  copy_files
  clasp push
  echo "Clasp project updated"
  cd ..
  rm -r clasp_uploads
  exit 1
fi

# Otherwise, only generate the files in clasp_uploads folder
cd clasp_uploads
copy_files