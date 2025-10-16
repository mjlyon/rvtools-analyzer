// RV Tools analyzer
// Ideally desinged to upload, and compare RVtools exports
// Designed to run as a web-app

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const express = require('express');

class RVToolsAnalyzer {
    constructor() {
        this.data = {};
        this.app = express();
        this.port = process.env.PORT || 3000;
        this.setupExpress();
    }

    setupExpress() {
        // Serve static files
        this.app.use(express.static(path.join(__dirname, 'public')));
        this.app.use(express.json({ limit: '50mb' }));
        this.app.use(express.urlencoded({ extended: true, limit: '50mb' }));
        
        // Root route - serve the dashboard
        this.app.get('/', (req, res) => {
            res.send(this.generateDashboardHTML());
        });

        // API routes
        this.app.post('/api/analyze', (req, res) => {
            try {
                const { fileData, fileName } = req.body;
                const analysis = this.analyzeData(fileData, fileName);
                res.json({ success: true, analysis });
            } catch (error) {
                console.error('Analysis error:', error);
                res.status(500).json({ success: false, error: error.message });
            }
        });

        this.app.post('/api/compare', (req, res) => {
            try {
                const { files } = req.body;
                const comparison = this.compareData(files);
                res.json({ success: true, comparison });
            } catch (error) {
                console.error('Comparison error:', error);
                res.status(500).json({ success: false, error: error.message });
            }
        });
    }

    // Load and parse RVTools Excel data
    parseExcelData(buffer) {
        const workbook = XLSX.read(buffer, { type: 'buffer' });
        const data = {};

        // Common RVTools sheets
        const sheets = ['vInfo', 'vHost', 'vDatastore', 'vCluster'];
        
        sheets.forEach(sheetName => {
            if (workbook.SheetNames.includes(sheetName)) {
                data[sheetName] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
            }
        });

        return data;
    }

    // Analyze VM statistics
    analyzeVMs(vInfo) {
        if (!vInfo || vInfo.length === 0) return null;

        const vms = vInfo.map(vm => ({
            name: vm['VM'] || vm['Name'] || 'Unknown',
            host: vm['Host'] || vm['ESX Host'] || 'N/A',
            cluster: vm['Cluster'] || 'N/A',
            cpus: parseInt(vm['CPUs'] || vm['Num CPUs'] || vm['vCPU'] || 0),
            memory: parseFloat(vm['Memory'] || vm['Memory MB'] || 0) / 1024,
            provisioned: parseFloat(vm['Provisioned MB'] || vm['Provisioned Space'] || vm['Provisioned'] || 0) / 1024,
            used: parseFloat(vm['In Use MB'] || vm['Used Space MB'] || vm['Used'] || 0) / 1024,
            powerState: vm['Powerstate'] || vm['Power State'] || 'Unknown'
        }));

        return {
            total: vms.length,
            poweredOn: vms.filter(vm => vm.powerState === 'poweredOn').length,
            poweredOff: vms.filter(vm => vm.powerState === 'poweredOff').length,
            totalvCPUs: vms.reduce((sum, vm) => sum + vm.cpus, 0),
            totalMemory: vms.reduce((sum, vm) => sum + vm.memory, 0),
            totalProvisioned: vms.reduce((sum, vm) => sum + vm.provisioned, 0),
            totalUsed: vms.reduce((sum, vm) => sum + vm.used, 0),
            avgCpus: vms.length > 0 ? vms.reduce((sum, vm) => sum + vm.cpus, 0) / vms.length : 0,
            avgMemory: vms.length > 0 ? vms.reduce((sum, vm) => sum + vm.memory, 0) / vms.length : 0,
            storageEfficiency: vms.reduce((sum, vm) => sum + vm.provisioned, 0) > 0 ? 
                (vms.reduce((sum, vm) => sum + vm.used, 0) / vms.reduce((sum, vm) => sum + vm.provisioned, 0)) * 100 : 0,
            vms: vms.slice(0, 100) // Limit for performance
        };
    }

    // Analyze Host statistics
    analyzeHosts(vHost) {
        if (!vHost || vHost.length === 0) return null;

        const hosts = vHost.map(host => ({
            name: host['Host'] || host['Hostname'] || host['ESX Host'] || 'Unknown',
            cluster: host['Cluster'] || 'N/A',
            cpuCores: parseInt(host['# CPU'] || host['CPU Cores'] || host['Num CPU'] || 0),
            memory: parseFloat(host['Memory'] || host['Memory GB'] || 0),
            vms: parseInt(host['# VMs'] || host['VM Count'] || host['VMs'] || 0),
            status: host['Status'] || host['Connection State'] || 'Unknown'
        }));

        return {
            total: hosts.length,
            connected: hosts.filter(h => h.status === 'Connected' || h.status === 'connected').length,
            totalCores: hosts.reduce((sum, h) => sum + h.cpuCores, 0),
            totalMemory: hosts.reduce((sum, h) => sum + h.memory, 0),
            totalVMs: hosts.reduce((sum, h) => sum + h.vms, 0),
            avgVMsPerHost: hosts.length > 0 ? hosts.reduce((sum, h) => sum + h.vms, 0) / hosts.length : 0,
            hosts: hosts
        };
    }

    // Analyze Storage statistics
    analyzeStorage(vDatastore) {
        if (!vDatastore || vDatastore.length === 0) return null;

        const datastores = vDatastore.map(ds => ({
            name: ds['Datastore'] || ds['Name'] || 'Unknown',
            type: ds['Type'] || 'Unknown',
            capacity: parseFloat(ds['Capacity GB'] || ds['Capacity'] || 0),
            used: parseFloat(ds['In Use GB'] || ds['In Use'] || ds['Used GB'] || 0),
            free: parseFloat(ds['Free GB'] || ds['Free'] || 0),
            vms: parseInt(ds['# VMs'] || ds['VM Count'] || ds['VMs'] || 0)
        }));

        const totalCapacity = datastores.reduce((sum, ds) => sum + ds.capacity, 0);
        const totalUsed = datastores.reduce((sum, ds) => sum + ds.used, 0);

        return {
            total: datastores.length,
            totalCapacity,
            totalUsed,
            totalFree: datastores.reduce((sum, ds) => sum + ds.free, 0),
            utilizationPercent: totalCapacity > 0 ? (totalUsed / totalCapacity) * 100 : 0,
            datastores: datastores
        };
    }

    // Main analysis function
    analyzeData(data, fileName) {
        const analysis = {
            fileName: fileName || 'Unknown',
            timestamp: new Date().toISOString(),
            vms: this.analyzeVMs(data.vInfo),
            hosts: this.analyzeHosts(data.vHost),
            storage: this.analyzeStorage(data.vDatastore)
        };

        return analysis;
    }

    // Compare multiple files
    compareData(files) {
        const analyses = files.map(file => this.analyzeData(file.data, file.name));
        
        return {
            files: analyses.map(a => a.fileName),
            timestamp: new Date().toISOString(),
            analyses: analyses
        };
    }

    // Generate dashboard HTML
    generateDashboardHTML() {
        return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>RVTools Sizing Analyzer</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/3.9.1/chart.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            margin: 0;
            padding: 20px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
        }
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 20px;
            padding: 30px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
        }
        h1 {
            color: #2c3e50;
            text-align: center;
            margin-bottom: 30px;
            font-size: 2.5em;
            font-weight: 300;
        }
        .upload-area {
            border: 3px dashed #667eea;
            border-radius: 15px;
            padding: 40px;
            text-align: center;
            margin-bottom: 30px;
            background: rgba(102, 126, 234, 0.1);
            transition: all 0.3s ease;
        }
        .upload-area:hover {
            border-color: #764ba2;
            background: rgba(102, 126, 234, 0.2);
        }
        .upload-button {
            background: linear-gradient(45deg, #667eea, #764ba2);
            color: white;
            border: none;
            padding: 15px 30px;
            border-radius: 25px;
            font-size: 16px;
            cursor: pointer;
            margin: 10px;
            transition: all 0.3s ease;
        }
        .upload-button:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 25px rgba(102, 126, 234, 0.6);
        }
        .file-list {
            margin-top: 20px;
            text-align: left;
        }
        .file-item {
            background: rgba(102, 126, 234, 0.1);
            padding: 10px 15px;
            margin: 5px 0;
            border-radius: 8px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .results {
            margin-top: 30px;
            padding: 20px;
            background: #f8f9fa;
            border-radius: 10px;
            display: none;
        }
        .metric-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin: 20px 0;
        }
        .metric-card {
            background: white;
            border-radius: 10px;
            padding: 20px;
            text-align: center;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        .metric-value {
            font-size: 2em;
            font-weight: bold;
            color: #667eea;
        }
        .metric-label {
            color: #666;
            margin-top: 5px;
        }
        .chart-container {
            background: white;
            border-radius: 10px;
            padding: 20px;
            margin: 20px 0;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        .loading {
            text-align: center;
            color: #667eea;
            font-size: 18px;
            margin: 20px 0;
        }
        .error {
            background: #fee;
            color: #c33;
            padding: 15px;
            border-radius: 8px;
            margin: 10px 0;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>🖥️ RVTools Sizing Analyzer</h1>
        
        <div class="upload-area" id="uploadArea">
            <h3>Upload RVTools Excel Files</h3>
            <p>Drop files here or click to browse</p>
            <input type="file" id="fileInput" multiple accept=".xlsx,.xls" style="display: none;">
            <button class="upload-button" onclick="document.getElementById('fileInput').click()">
                Choose Files
            </button>
            <div class="file-list" id="fileList"></div>
        </div>

        <button class="upload-button" onclick="analyzeFiles()" id="analyzeBtn" style="display: none;">
            Analyze Files
        </button>
        <button class="upload-button" onclick="compareFiles()" id="compareBtn" style="display: none;">
            Compare Files
        </button>

        <div id="loading" class="loading" style="display: none;">
            Processing files... ⚙️
        </div>

        <div id="error" class="error" style="display: none;"></div>

        <div id="results" class="results"></div>
    </div>

    <script>
        let uploadedFiles = [];

        document.getElementById('fileInput').addEventListener('change', handleFileSelect);
        
        // Drag and drop functionality
        const uploadArea = document.getElementById('uploadArea');
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.style.borderColor = '#28a745';
        });
        
        uploadArea.addEventListener('dragleave', (e) => {
            e.preventDefault();
            uploadArea.style.borderColor = '#667eea';
        });
        
        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.style.borderColor = '#667eea';
            handleFileSelect({target: {files: e.dataTransfer.files}});
        });

        function handleFileSelect(event) {
            const files = Array.from(event.target.files);
            uploadedFiles = files.filter(file => file.name.match(/\\.(xlsx|xls)$/i));
            
            displayFileList();
            
            if (uploadedFiles.length > 0) {
                document.getElementById('analyzeBtn').style.display = 'inline-block';
                if (uploadedFiles.length > 1) {
                    document.getElementById('compareBtn').style.display = 'inline-block';
                }
            }
        }

        function displayFileList() {
            const fileList = document.getElementById('fileList');
            fileList.innerHTML = uploadedFiles.map((file, index) => 
                '<div class="file-item">' +
                    '<span>' + file.name + ' (' + (file.size/1024/1024).toFixed(2) + ' MB)</span>' +
                    '<button onclick="removeFile(' + index + ')" style="background: #e74c3c; color: white; border: none; padding: 5px 10px; border-radius: 5px; cursor: pointer;">Remove</button>' +
                '</div>'
            ).join('');
        }

        function removeFile(index) {
            uploadedFiles.splice(index, 1);
            displayFileList();
            
            if (uploadedFiles.length === 0) {
                document.getElementById('analyzeBtn').style.display = 'none';
                document.getElementById('compareBtn').style.display = 'none';
            } else if (uploadedFiles.length === 1) {
                document.getElementById('compareBtn').style.display = 'none';
            }
        }

        async function analyzeFiles() {
            if (uploadedFiles.length === 0) return;
            
            showLoading(true);
            hideError();
            
            try {
                const file = uploadedFiles[0];
                const fileData = await readFileAsArrayBuffer(file);
                const workbook = XLSX.read(fileData, { type: 'array' });
                
                const data = {};
                const sheets = ['vInfo', 'vHost', 'vDatastore', 'vCluster'];
                sheets.forEach(sheetName => {
                    if (workbook.SheetNames.includes(sheetName)) {
                        data[sheetName] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
                    }
                });
                
                const response = await fetch('/api/analyze', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ fileData: data, fileName: file.name })
                });
                
                const result = await response.json();
                
                if (result.success) {
                    displayAnalysis(result.analysis);
                } else {
                    showError('Analysis failed: ' + result.error);
                }
            } catch (error) {
                showError('Error processing file: ' + error.message);
            } finally {
                showLoading(false);
            }
        }

        async function compareFiles() {
            if (uploadedFiles.length < 2) return;
            
            showLoading(true);
            hideError();
            
            try {
                const files = [];
                
                for (const file of uploadedFiles) {
                    const fileData = await readFileAsArrayBuffer(file);
                    const workbook = XLSX.read(fileData, { type: 'array' });
                    
                    const data = {};
                    const sheets = ['vInfo', 'vHost', 'vDatastore', 'vCluster'];
                    sheets.forEach(sheetName => {
                        if (workbook.SheetNames.includes(sheetName)) {
                            data[sheetName] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
                        }
                    });
                    
                    files.push({ name: file.name, data: data });
                }
                
                const response = await fetch('/api/compare', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ files: files })
                });
                
                const result = await response.json();
                
                if (result.success) {
                    displayComparison(result.comparison);
                } else {
                    showError('Comparison failed: ' + result.error);
                }
            } catch (error) {
                showError('Error comparing files: ' + error.message);
            } finally {
                showLoading(false);
            }
        }

        function readFileAsArrayBuffer(file) {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = e => resolve(new Uint8Array(e.target.result));
                reader.onerror = reject;
                reader.readAsArrayBuffer(file);
            });
        }

        function displayAnalysis(analysis) {
            const results = document.getElementById('results');
            results.style.display = 'block';
            
            let html = '<h2>📊 Analysis Results: ' + analysis.fileName + '</h2>';
            html += '<div class="metric-grid">';
            
            if (analysis.vms) {
                html += '<div class="metric-card"><div class="metric-value">' + analysis.vms.total + '</div><div class="metric-label">Total VMs</div></div>';
                html += '<div class="metric-card"><div class="metric-value">' + analysis.vms.poweredOn + '</div><div class="metric-label">Powered On</div></div>';
                html += '<div class="metric-card"><div class="metric-value">' + analysis.vms.totalvCPUs + '</div><div class="metric-label">Total vCPUs</div></div>';
                html += '<div class="metric-card"><div class="metric-value">' + analysis.vms.totalMemory.toFixed(0) + ' GB</div><div class="metric-label">Total Memory</div></div>';
            }
            
            if (analysis.hosts) {
                html += '<div class="metric-card"><div class="metric-value">' + analysis.hosts.total + '</div><div class="metric-label">Total Hosts</div></div>';
                html += '<div class="metric-card"><div class="metric-value">' + analysis.hosts.totalCores + '</div><div class="metric-label">CPU Cores</div></div>';
            }
            
            if (analysis.storage) {
                html += '<div class="metric-card"><div class="metric-value">' + (analysis.storage.totalCapacity / 1024).toFixed(1) + ' TB</div><div class="metric-label">Storage Capacity</div></div>';
                html += '<div class="metric-card"><div class="metric-value">' + analysis.storage.utilizationPercent.toFixed(1) + '%</div><div class="metric-label">Storage Used</div></div>';
            }
            
            html += '</div>';
            
            results.innerHTML = html;
        }

        function displayComparison(comparison) {
            const results = document.getElementById('results');
            results.style.display = 'block';
            
            let html = '<h2>📈 Comparison Results</h2>';
            html += '<div style="overflow-x: auto;"><table style="width: 100%; border-collapse: collapse;">';
            html += '<tr style="background: #f8f9fa;"><th style="padding: 10px; border: 1px solid #ddd;">File</th><th style="padding: 10px; border: 1px solid #ddd;">VMs</th><th style="padding: 10px; border: 1px solid #ddd;">Hosts</th><th style="padding: 10px; border: 1px solid #ddd;">Storage (TB)</th></tr>';
            
            comparison.analyses.forEach(analysis => {
                html += '<tr>';
                html += '<td style="padding: 10px; border: 1px solid #ddd;">' + analysis.fileName + '</td>';
                html += '<td style="padding: 10px; border: 1px solid #ddd;">' + (analysis.vms ? analysis.vms.total : 0) + '</td>';
                html += '<td style="padding: 10px; border: 1px solid #ddd;">' + (analysis.hosts ? analysis.hosts.total : 0) + '</td>';
                html += '<td style="padding: 10px; border: 1px solid #ddd;">' + (analysis.storage ? (analysis.storage.totalCapacity / 1024).toFixed(1) : 0) + '</td>';
                html += '</tr>';
            });
            
            html += '</table></div>';
            
            results.innerHTML = html;
        }

        function showLoading(show) {
            document.getElementById('loading').style.display = show ? 'block' : 'none';
        }

        function showError(message) {
            const errorDiv = document.getElementById('error');
            errorDiv.textContent = message;
            errorDiv.style.display = 'block';
        }

        function hideError() {
            document.getElementById('error').style.display = 'none';
        }
    </script>
</body>
</html>`;
    }

    // Command line analysis
    analyzeFile(filePath) {
        if (!fs.existsSync(filePath)) {
            throw new Error(`File not found: ${filePath}`);
        }

        console.log(`Loading file: ${filePath}`);
        const workbook = XLSX.readFile(filePath);
        const data = {};

        const sheets = ['vInfo', 'vHost', 'vDatastore', 'vCluster'];
        sheets.forEach(sheetName => {
            if (workbook.SheetNames.includes(sheetName)) {
                data[sheetName] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
                console.log(`  - Loaded ${data[sheetName].length} rows from ${sheetName}`);
            }
        });

        const analysis = this.analyzeData(data, path.basename(filePath));
        this.printConsoleReport(analysis);
        
        // Save reports
        const baseName = path.basename(filePath, path.extname(filePath));
        fs.writeFileSync(`${baseName}_analysis.json`, JSON.stringify(analysis, null, 2));
        
        console.log(`\n📄 Report saved: ${baseName}_analysis.json`);
        return analysis;
    }

    printConsoleReport(analysis) {
        console.log('\n' + '='.repeat(60));
        console.log(`RVTools Analysis Report: ${analysis.fileName}`);
        console.log('='.repeat(60));

        if (analysis.vms) {
            console.log('\n📊 VIRTUAL MACHINES');
            console.log('-'.repeat(30));
            console.log(`Total VMs: ${analysis.vms.total.toLocaleString()}`);
            console.log(`Powered On: ${analysis.vms.poweredOn.toLocaleString()} (${((analysis.vms.poweredOn/analysis.vms.total)*100).toFixed(1)}%)`);
            console.log(`Total vCPUs: ${analysis.vms.totalvCPUs.toLocaleString()}`);
            console.log(`Total Memory: ${analysis.vms.totalMemory.toFixed(1)} GB`);
            console.log(`Avg vCPUs/VM: ${analysis.vms.avgCpus.toFixed(1)}`);
            console.log(`Avg Memory/VM: ${analysis.vms.avgMemory.toFixed(1)} GB`);
            console.log(`Storage Efficiency: ${analysis.vms.storageEfficiency.toFixed(1)}%`);
        }

        if (analysis.hosts) {
            console.log('\n🖥️  HOSTS');
            console.log('-'.repeat(30));
            console.log(`Total Hosts: ${analysis.hosts.total.toLocaleString()}`);
            console.log(`Connected: ${analysis.hosts.connected.toLocaleString()}`);
            console.log(`Total CPU Cores: ${analysis.hosts.totalCores.toLocaleString()}`);
            console.log(`Total Memory: ${analysis.hosts.totalMemory.toFixed(1)} GB`);
            console.log(`Avg VMs/Host: ${analysis.hosts.avgVMsPerHost.toFixed(1)}`);
        }

        if (analysis.storage) {
            console.log('\n💾 STORAGE');
            console.log('-'.repeat(30));
            console.log(`Total Datastores: ${analysis.storage.total.toLocaleString()}`);
            console.log(`Total Capacity: ${(analysis.storage.totalCapacity/1024).toFixed(1)} TB`);
            console.log(`Total Used: ${(analysis.storage.totalUsed/1024).toFixed(1)} TB`);
            console.log(`Utilization: ${analysis.storage.utilizationPercent.toFixed(1)}%`);
        }
    }

    // Start the web server
    start() {
        this.app.listen(this.port, '0.0.0.0', () => {
            console.log(`🚀 RVTools Analyzer running at http://localhost:${this.port}`);
            console.log(`   Access from network: http://YOUR_SERVER_IP:${this.port}`);
            console.log('\nUsage:');
            console.log('  Web Interface: Open the URL above in your browser');
            console.log('  Command Line: node rvtools-analyzer.js /path/to/file.xlsx');
        });
    }
}

// Main execution
function main() {
    const analyzer = new RVToolsAnalyzer();
    
    if (process.argv.length >= 3) {
        // Command line mode
        const filePath = process.argv[2];
        try {
            analyzer.analyzeFile(filePath);
        } catch (error) {
            console.error('❌ Error:', error.message);
            process.exit(1);
        }
    } else {
        // Web server mode
        analyzer.start();
    }
}

if (require.main === module) {
    main();
} else {
    module.exports = RVToolsAnalyzer;
}
