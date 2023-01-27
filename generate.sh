#!/bin/bash

audit_type=$1

# Check if the audit_type is valid
if [[ "$audit_type" != "app" && "$audit_type" != "web" && "$audit_type" != "ux" ]]; then
  echo "Error: Invalid audit type"
  exit 1
fi

# Create the clasp_uploads folder if it doesn't already exist
if [ ! -d "clasp_uploads" ]; then
  mkdir clasp_uploads
fi

cp "sls_core.js" "clasp_uploads/sls_core.js"

# Copy files based on audit_type value
if [ "$audit_type" == "app" ]; then
    cp "sls_app.js" "clasp_uploads/sls_app.js"
elif [ "$audit_type" == "web" ]; then
    cp "sls_web.js" "clasp_uploads/sls_web.js"
elif [ "$audit_type" == "ux" ]; then
    cp "sls_ux.js" "clasp_uploads/sls_ux.js"
fi

echo "Files successfully copied to clasp_uploads folder"