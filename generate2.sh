#!/bin/bash

 # @license
 # Copyright 2023 Google LLC
 #
 # Licensed under the Apache License, Version 2.0 (the "License");
 # you may not use this file except in compliance with the License.
 # You may obtain a copy of the License at
 
 #      http://www.apache.org/licenses/LICENSE-2.0
 
 # Unless required by applicable law or agreed to in writing, software
 # distributed under the License is distributed on an "AS IS" BASIS,
 # WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 # See the License for the specific language governing permissions and
 # limitations under the License.


audit_type=$1
clasp_upload=$2
clasp_script_url=$3

# Check if the audit_type is valid
if [[ "$audit_type" != "app" && "$audit_type" != "web" && "$audit_type" != "ux" && "$audit_type" != "sustainability" ]]; then
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
  cp -r ../src/* .
  rm -r audits
  cp ../src/audits/$audit_type.js audits/$audit_type.js
}

# Create the output folder if it doesn't already exist
if [ ! -d "output" ]; then
  mkdir output
fi

# Check if the user wants to upload to clasp
if [[ "$clasp_upload" == "--clasp-upload" ]]; then
  check_login
  if [ -z "$clasp_script_url" ]; then
    echo "Error: Script ID or URL not present"
    exit 1
  fi
  cd output
  clasp clone "$clasp_script_url"
  copy_files
  clasp push
  echo "Clasp project updated"
  cd ..
  rm -r output
  exit 1
fi

# Otherwise, only generate the files in output folder
cd output
copy_files