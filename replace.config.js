// This configuration tells 'replace-in-file' to modify your manifest.xml
// It replaces the token with https://localhost:3000
const config = {
  files: 'manifest.xml',
  from: /~remoteAppUrl~/g,
  to: 'https://localhost:3000',
};

// We need to export this configuration for the script to use it.
module.exports = config;

