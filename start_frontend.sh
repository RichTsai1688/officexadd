#!/bin/bash
cd "$(dirname "$0")/frontend"
npx http-server -p 3010 --cors
