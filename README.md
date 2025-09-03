# RVtools-analyzer
RV Tools Analyzer is designed to take RVTools via upload and provide some analysis of utilization
- Designed for portability
- Designed for local path or upload


## Step 1 - Install Node
### For Ubuntu/Debian
curl -fsSL https://deb.nodesource.com/setup_18.x | sudo -E bash -
sudo apt-get install -y nodejs

### For CentOS/RHEL/Fedora
curl -fsSL https://rpm.nodesource.com/setup_18.x | sudo bash -
sudo dnf install -y nodejs npm

### Verify installation
node --version
npm --version

## Step 2 - Create App Directory

sudo mkdir -p /opt/rvtools-analyzer
cd /opt/rvtools-analyzer

## Set ownership (replace 'username' with your user)
sudo chown -R $USER:$USER /opt/rvtools-analyzer
