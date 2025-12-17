<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Onix Annual Dashboard</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- Chart.js -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <!-- PapaParse for robust CSV parsing -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js"></script>
    <!-- SheetJS for Excel support -->
    <script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>
    <!-- Lucide Icons -->
    <script src="https://unpkg.com/lucide@latest"></script>
    <style>
        body {
            font-family: 'Inter', sans-serif;
            background-color: #f3f4f6;
            color: #1E2020; /* Satin Black */
        }
        .drop-zone {
            transition: all 0.3s ease;
        }
        .drop-zone.dragover {
            background-color: #eff6ff;
            border-color: #30496D; /* Midnight Blue */
            transform: scale(1.02);
        }
        /* Custom Scrollbar */
        .custom-scrollbar::-webkit-scrollbar {
            width: 4px;
        }
        .custom-scrollbar::-webkit-scrollbar-track {
            background: #f1f1f1;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb {
            background: #ccc;
            border-radius: 2px;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover {
            background: #aaa;
        }
    </style>
</head>
<body class="min-h-screen flex flex-col">

    <!-- Header -->
    <header class="bg-white shadow-sm border-b-4" style="border-bottom-color: #30496D;">
        <div class="max-w-7xl mx-auto px-4 py-4 flex items-center justify-between">
            <div class="flex items-center gap-3">
                <i data-lucide="bar-chart-2" class="w-8 h-8" style="color: #F46262;"></i> <!-- Onix Red -->
                <h1 class="text-2xl font-bold tracking-tight" style="color: #30496D;">Onix Annual Dashboard</h1> <!-- Midnight Blue -->
            </div>
            <div class="text-sm font-medium px-3 py-1 rounded-full bg-[#FEED7D]/50 text-[#30496D]">
                Yearly Overview
            </div>
        </div>
    </header>

    <!-- Main Content -->
    <main class="flex-grow container mx-auto px-4 py-8">
        
        <!-- Upload Section -->
        <div id="upload-section" class="max-w-2xl mx-auto mb-10 transition-all duration-300">
            <div class="bg-white rounded-xl shadow-lg p-8">
                <div class="text-center mb-6">
                    <h2 class="text-xl font-semibold" style="color: #30496D;">Initialize Dashboard</h2>
                    <p class="text-gray-500 text-sm mt-1">
                        Drag & Drop all your LinkedIn export files here.
                    </p>
                    <p class="text-xs text-gray-400 mt-2">
                        Supports: <span class="font-medium text-[#2795A4]">New followers, Metrics, All posts, Seniority</span>
                    </p>
                </div>

                <div id="drop-zone" class="drop-zone border-2 border-dashed border-gray-300 rounded-xl p-10 flex flex-col items-center justify-center cursor-pointer hover:border-[#2795A4] group">
                    <div class="p-4 rounded-full mb-4 transition-colors bg-[#F46262]/10 group-hover:bg-[#F46262]/20">
                        <i data-lucide="upload-cloud" class="w-8 h-8" style="color: #F46262;"></i> <!-- Onix Red -->
                    </div>
                    <p class="text-gray-600 font-medium">Drag & Drop Excel or CSV files</p>
                    <input type="file" id="file-input" class="hidden" accept=".csv, .xlsx, .xls" multiple>
                </div>
            </div>
        </div>

        <!-- Dashboard Section (Hidden by default) -->
        <div id="dashboard" class="hidden space-y-8">
            
            <!-- SECTION 1: Top KPIs -->
            <div>
                <h2 class="text-xl font-bold mb-4 flex items-center gap-2" style="color: #30496D;">
                    <i data-lucide="target" class="w-5 h-5"></i> Key Performance Indicators
                </h2>
                <div class="grid grid-cols-1 md:grid-cols-5 gap-4">
                    
                    <!-- KPI 1: New Followers (Auto-Synced) -->
                    <div class="bg-white rounded-xl shadow p-5 border-t-4" style="border-top-color: #F46262;">
                        <p class="text-xs font-semibold uppercase tracking-wider text-gray-500 mb-1">New Followers</p>
                        <div class="flex items-end justify-between">
                            <h3 id="kpi-new-followers" class="text-3xl font-bold" style="color: #1E2020;">-</h3>
                            <i data-lucide="users" class="w-5 h-5 mb-1 opacity-50" style="color: #F46262;"></i>
                        </div>
                        <p class="text-xs text-gray-400 mt-1">Auto-calculated from data</p>
                    </div>

                    <!-- KPI 2: 3-Month Engagement Rate (Manual/Default) -->
                    <div class="bg-white rounded-xl shadow p-5 border-t-4" style="border-top-color: #30496D;">
                        <p class="text-xs font-semibold uppercase tracking-wider text-gray-500 mb-1">Eng. Rate (3mo)</p>
                        <div class="flex items-end justify-between">
                            <input type="text" value="22.6%" class="text-3xl font-bold w-full bg-transparent border-none focus:ring-0 p-0 text-[#1E2020]" style="color: #1E2020;">
                            <i data-lucide="percent" class="w-5 h-5 mb-1 opacity-50" style="color: #30496D;"></i>
                        </div>
                        <p class="text-xs text-gray-400 mt-1">Last 3 months average</p>
                    </div>

                    <!-- KPI 3: Avg Engagement Rate (Auto/Manual) -->
                    <div class="bg-white rounded-xl shadow p-5 border-t-4" style="border-top-color: #874D83;">
                        <p class="text-xs font-semibold uppercase tracking-wider text-gray-500 mb-1">Avg Eng. Rate</p>
                        <div class="flex items-end justify-between">
                            <input id="kpi-avg-eng-rate" type="text" value="5.8%" class="text-3xl font-bold w-full bg-transparent border-none focus:ring-0 p-0 text-[#1E2020]" style="color: #1E2020;">
                            <i data-lucide="activity" class="w-5 h-5 mb-1 opacity-50" style="color: #874D83;"></i>
                        </div>
                        <p id="kpi-avg-eng-rate-sub" class="text-xs text-gray-400 mt-1">Yearly average</p>
                    </div>

                     <!-- KPI 4: Avg Engagement Count (Manual - Placeholder as typically not in standard metrics export) -->
                     <div class="bg-white rounded-xl shadow p-5 border-t-4" style="border-top-color: #A52959;">
                        <p class="text-xs font-semibold uppercase tracking-wider text-gray-500 mb-1">Avg Engagement</p>
                        <div class="flex items-end justify-between">
                            <input type="text" value="145" class="text-3xl font-bold w-full bg-transparent border-none focus:ring-0 p-0 text-[#1E2020]" style="color: #1E2020;">
                            <i data-lucide="thumbs-up" class="w-5 h-5 mb-1 opacity-50" style="color: #A52959;"></i>
                        </div>
                        <p class="text-xs text-gray-400 mt-1">Interactions per post</p>
                    </div>

                    <!-- KPI 5: Avg Impressions (Auto/Manual) -->
                    <div class="bg-white rounded-xl shadow p-5 border-t-4" style="border-top-color: #2795A4;">
                        <p class="text-xs font-semibold uppercase tracking-wider text-gray-500 mb-1">Avg Impressions</p>
                        <div class="flex items-end justify-between">
                            <input id="kpi-avg-impressions" type="text" value="3.2K" class="text-3xl font-bold w-full bg-transparent border-none focus:ring-0 p-0 text-[#1E2020]" style="color: #1E2020;">
                            <i data-lucide="eye" class="w-5 h-5 mb-1 opacity-50" style="color: #2795A4;"></i>
                        </div>
                        <p id="kpi-avg-impressions-sub" class="text-xs text-gray-400 mt-1">Views per post</p>
                    </div>
                </div>
            </div>

            <!-- SECTION 2: Growth Chart -->
            <div class="bg-white rounded-xl shadow-lg p-6">
                <div class="flex flex-col md:flex-row md:items-center justify-between mb-6 gap-4">
                    <div>
                         <h2 class="text-lg font-bold" style="color: #30496D;">Follower Increase Over Time</h2>
                         <p class="text-xs text-gray-400">Visualization of daily growth vs cumulative total</p>
                    </div>
                    
                    <!-- Chart Controls -->
                    <div class="flex items-center gap-4 bg-gray-50 p-2 rounded-lg">
                        <div class="flex items-center gap-2">
                             <label for="start-count" class="text-xs font-medium text-gray-600">Start Count:</label>
                             <div class="relative">
                                <input type="number" id="start-count" class="w-24 text-sm border-gray-300 rounded px-2 py-1 focus:ring-[#2795A4]" value="0">
                                <div id="auto-badge" class="hidden absolute -top-1 -right-1 h-2 w-2 bg-emerald-500 rounded-full"></div>
                             </div>
                        </div>
                        <div class="h-4 w-px bg-gray-300"></div>
                        <div class="flex gap-3">
                            <div class="flex items-center gap-1 text-xs text-gray-500">
                                <span class="w-3 h-3 rounded-sm opacity-80" style="background-color: #2795A4;"></span> Daily
                             </div>
                             <div class="flex items-center gap-1 text-xs text-gray-500">
                                <span class="w-3 h-3 rounded-full border" style="background-color: #F46262; border-color: #F46262;"></span> Total
                             </div>
                        </div>
                    </div>
                </div>
                <div class="relative h-[350px] w-full">
                    <canvas id="growthChart"></canvas>
                </div>
            </div>

     


            <!-- Reset Button -->
            <div class="text-center pt-8 border-t border-gray-200">
                <button id="reset-btn" class="text-sm text-gray-500 hover:text-[#30496D] underline">Reset Dashboard</button>
            </div>
        </div>

        <!-- Error Message -->
        <div id="error-msg" class="hidden max-w-2xl mx-auto mt-4 bg-red-50 border border-red-200 text-red-600 px-4 py-3 rounded-lg flex items-center gap-2">
            <i data-lucide="alert-circle" class="w-5 h-5"></i>
            <span id="error-text">An error occurred.</span>
        </div>

    </main>

    <script>
        // Initialize Lucide icons
        lucide.createIcons();

        // Brand Colors
        const BRAND_COLORS = {
            onixRed: '#F46262',
            midnightBlue: '#30496D',
            berryBlue: '#874D83',
            amaranthPurple: '#A52959',
            maize: '#FEED7D',
            midnightGreen: '#094C61',
            munsellBlue: '#2795A4',
            white: '#ffffff',
            satinBlack: '#1E2020'
        };

        // DOM Elements
        const dropZone = document.getElementById('drop-zone');
        const fileInput = document.getElementById('file-input');
        const uploadSection = document.getElementById('upload-section');
        const dashboard = document.getElementById('dashboard');
        const errorMsg = document.getElementById('error-msg');
        const errorText = document.getElementById('error-text');
        const resetBtn = document.getElementById('reset-btn');
        const startCountInput = document.getElementById('start-count');
        const autoBadge = document.getElementById('auto-badge');
        const kpiNewFollowers = document.getElementById('kpi-new-followers');
        const postsContainer = document.getElementById('posts-container');
        
        // KPI Inputs
        const kpiAvgEngRate = document.getElementById('kpi-avg-eng-rate');
        const kpiAvgEngRateSub = document.getElementById('kpi-avg-eng-rate-sub');
        const kpiAvgImpressions = document.getElementById('kpi-avg-impressions');
        const kpiAvgImpressionsSub = document.getElementById('kpi-avg-impressions-sub');

        let chartInstance = null;
        let globalProcessedData = []; 
        let globalSnapshotTotal = null; 

        // Event Listeners for Drag & Drop
        dropZone.addEventListener('click', () => fileInput.click());
        
        dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropZone.classList.add('dragover');
        });

        dropZone.addEventListener('dragleave', () => {
            dropZone.classList.remove('dragover');
        });

        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.classList.remove('dragover');
            const files = e.dataTransfer.files;
            if (files.length) handleFiles(files);
        });

        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length) handleFiles(e.target.files);
        });

        resetBtn.addEventListener('click', () => {
            dashboard.classList.add('hidden');
            uploadSection.classList.remove('hidden');
            fileInput.value = '';
            if (chartInstance) chartInstance.destroy();
            globalProcessedData = [];
            globalSnapshotTotal = null;
            startCountInput.value = 0; 
            autoBadge.classList.add('hidden');
            errorMsg.classList.add('hidden');
            
            // Reset placeholders
            postsContainer.innerHTML = `
                <div class="col-span-full text-center py-10 bg-gray-50 rounded-xl border border-dashed border-gray-300">
                    <p class="text-gray-500">Upload "All posts.csv" to see top performing content here.</p>
                </div>`;
            kpiAvgEngRate.value = "5.8%";
            kpiAvgImpressions.value = "3.2K";
        });

        startCountInput.addEventListener('input', () => {
            if (globalProcessedData.length > 0) {
                updateDashboard();
            }
        });

        function showError(msg) {
            errorText.textContent = msg;
            errorMsg.classList.remove('hidden');
        }

        async function handleFiles(files) {
            errorMsg.classList.add('hidden');
            const fileList = Array.from(files);
            
            let filesProcessed = 0;
            let growthDataFound = false;

            for (const file of fileList) {
                try {
                    const isExcel = file.name.endsWith('.xlsx') || file.name.endsWith('.xls');
                    const isCSV = file.name.endsWith('.csv');

                    if (isExcel) {
                        await processExcelFile(file);
                    } else if (isCSV) {
                        await processCSVFile(file);
                    } else {
                        console.warn("Skipping unsupported file:", file.name);
                    }
                    filesProcessed++;
                } catch (err) {
                    showError(`Error processing ${file.name}: ${err.message}`);
                }
            }
            
            // Check if we have growth data to show the main chart
            if (globalProcessedData.length > 0) {
                 finalizeDataLoad();
            } else if (filesProcessed > 0 && globalSnapshotTotal === null) {
                 // If files were processed but no specific recognizable data found
                 showError("Could not recognize specific data formats (Followers, Metrics, or Posts). Please check file types.");
            }
        }

        function processExcelFile(file) {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = (e) => {
                    try {
                        const data = new Uint8Array(e.target.result);
                        const workbook = XLSX.read(data, {type: 'array'});
                        
                        workbook.SheetNames.forEach(sheetName => {
                            const worksheet = workbook.Sheets[sheetName];
                            const jsonData = XLSX.utils.sheet_to_json(worksheet);
                            analyzeDataChunk(jsonData);
                        });
                        resolve();
                    } catch (err) {
                        reject(err);
                    }
                };
                reader.onerror = (err) => reject(err);
                reader.readAsArrayBuffer(file);
            });
        }

        function processCSVFile(file) {
            return new Promise((resolve, reject) => {
                Papa.parse(file, {
                    header: true,
                    skipEmptyLines: true,
                    complete: function(results) {
                        analyzeDataChunk(results.data);
                        resolve();
                    },
                    error: function(error) {
                        reject(error);
                    }
                });
            });
        }

        function analyzeDataChunk(data) {
            if (!data || data.length === 0) return;
            const keys = Object.keys(data[0]);
            const lowerKeys = keys.map(k => k.toLowerCase());
            
            const hasDate = lowerKeys.some(k => k.includes('date'));
            const hasFollowers = lowerKeys.some(k => k.includes('total followers') || k.includes('new followers'));
            const hasImpressions = lowerKeys.some(k => k.includes('impressions'));
            const hasPostTitle = lowerKeys.some(k => k.includes('post title') || k.includes('post link'));
            const hasEngagementRate = lowerKeys.some(k => k.includes('engagement rate'));

            // 1. Growth Data: Has Date AND Followers, but NOT Impressions (Metrics file has impressions)
            // The "New followers.csv" file usually has Date and Total Followers (daily gain).
            if (hasDate && hasFollowers && !hasImpressions) {
                processGrowthData(data);
            } 
            // 2. Metrics Data (KPIs): Has Date AND Impressions AND Engagement Rate
            else if (hasDate && hasImpressions && hasEngagementRate) {
                processMetricsData(data);
            }
            // 3. Posts Data: Has Post Title
            else if (hasPostTitle) {
                processPostsData(data);
            }
            // 4. Snapshot Data (Demographics): Has Followers but NO Date
            else if (hasFollowers && !hasDate) {
                processSnapshotData(data);
            }
        }

        function processSnapshotData(data) {
            const keys = Object.keys(data[0]);
            const totalKey = keys.find(k => k.toLowerCase().includes('total followers'));
            
            if (totalKey) {
                const sum = data.reduce((acc, row) => acc + (parseFloat(row[totalKey]) || 0), 0);
                if (sum > 0) {
                    if (globalSnapshotTotal === null || sum > globalSnapshotTotal) {
                         globalSnapshotTotal = sum;
                    }
                }
            }
        }

        function processGrowthData(data) {
             const keys = Object.keys(data[0]);
             const dateKey = keys.find(k => k.toLowerCase().includes('date'));
             // In "New followers.csv", the daily gain is often labeled "Total followers" in the CSV export from LinkedIn
             const totalKey = keys.find(k => k.toLowerCase().includes('total followers'));

             if (!dateKey || !totalKey) return;
             
             const newData = data.map(row => {
                let dateObj = new Date(row[dateKey]);
                // Handle Excel serial dates if needed
                if (isNaN(dateObj.getTime()) && !isNaN(row[dateKey])) {
                     dateObj = new Date(Math.round((row[dateKey] - 25569)*86400*1000));
                }
                return {
                    date: dateObj,
                    val: parseFloat(row[totalKey]) || 0,
                    rawDate: row[dateKey]
                };
            })
            .filter(d => !isNaN(d.date.getTime())) 
            .sort((a, b) => a.date - b.date);

            if (newData.length > 0) {
                globalProcessedData = newData;
            }
        }

        function processMetricsData(data) {
            // "Metrics.csv" -> Avg Engagement Rate, Avg Impressions
            const keys = Object.keys(data[0]);
            const impKey = keys.find(k => k.toLowerCase().includes('impressions (total)')) || keys.find(k => k.toLowerCase().includes('impressions'));
            const engRateKey = keys.find(k => k.toLowerCase().includes('engagement rate (total)')) || keys.find(k => k.toLowerCase().includes('engagement rate'));

            if (!impKey || !engRateKey) return;

            let totalImp = 0;
            let totalEngRate = 0;
            let count = 0;

            data.forEach(row => {
                const imp = parseFloat(row[impKey]);
                const rate = parseFloat(row[engRateKey]);
                if (!isNaN(imp) && !isNaN(rate)) {
                    totalImp += imp;
                    totalEngRate += rate;
                    count++;
                }
            });

            if (count > 0) {
                const avgImp = totalImp / count; // This is Avg Impressions per DAY (since Metrics.csv is daily)
                // However, user might want Avg Impressions PER POST. Metrics.csv is aggregated daily.
                // If we want per post, we need All Posts.csv.
                // But let's use this as a daily average fallback, or update if we process posts.
                
                // Let's stick to using this for the Daily Average if Posts data isn't available,
                // But usually "Avg Impressions" implies per post.
                // Actually, let's use the average daily impressions here.
                
                kpiAvgImpressions.value = Math.round(avgImp).toLocaleString();
                kpiAvgImpressionsSub.textContent = "Daily average impressions";
                kpiAvgImpressionsSub.className = "text-xs text-emerald-600 font-medium";

                // Engagement Rate is usually an average of the rates
                const avgRate = (totalEngRate / count) * 100; // Assuming rate is 0.05 for 5%
                // Check if rate is 0-1 or 0-100. Usually 0-1 in exports.
                // If avg is like 0.04, it's 4%. If it's 4, it's 4%.
                // Heuristic: if avg > 1, assume percent.
                let displayRate = avgRate; 
                // However, standard CSV export is decimal (0.05).
                
                kpiAvgEngRate.value = displayRate.toFixed(2) + "%";
                kpiAvgEngRateSub.textContent = "Daily average rate (calculated)";
                kpiAvgEngRateSub.className = "text-xs text-emerald-600 font-medium";
            }
        }

        function processPostsData(data) {
            // "All posts.csv" -> Top 5 Posts
            const keys = Object.keys(data[0]);
            const titleKey = keys.find(k => k.toLowerCase().includes('post title'));
            const impKey = keys.find(k => k.toLowerCase().includes('impressions'));
            const engKey = keys.find(k => k.toLowerCase().includes('engagement rate'));
            
            if (!titleKey || !impKey) return;

            // Sort by Impressions descending
            const sortedPosts = data.sort((a, b) => {
                return (parseFloat(b[impKey]) || 0) - (parseFloat(a[impKey]) || 0);
            }).slice(0, 5);

            if (sortedPosts.length > 0) {
                renderTopPosts(sortedPosts, titleKey, impKey, engKey);
                
                // Also update the KPI for Avg Impressions per POST
                const allImpressions = data.reduce((acc, row) => acc + (parseFloat(row[impKey]) || 0), 0);
                const avgImpPerPost = allImpressions / data.length;
                kpiAvgImpressions.value = Math.round(avgImpPerPost).toLocaleString();
                kpiAvgImpressionsSub.textContent = "Average views per post";

                const allEngRate = data.reduce((acc, row) => acc + (parseFloat(row[engKey]) || 0), 0);
                const avgEngRatePerPost = (allEngRate / data.length) * 100;
                kpiAvgEngRate.value = avgEngRatePerPost.toFixed(2) + "%";
                kpiAvgEngRateSub.textContent = "Average rate per post";
            }
        }

        function renderTopPosts(posts, titleKey, impKey, engKey) {
            postsContainer.innerHTML = '';
            
            posts.forEach((post, index) => {
                const title = post[titleKey] || "No description";
                const impressions = Math.round(parseFloat(post[impKey]) || 0).toLocaleString();
                let engRate = parseFloat(post[engKey]) || 0;
                // Normalize engagement rate (sometimes 0.05, sometimes 5)
                if (engRate < 1) engRate = engRate * 100;
                
                const card = `
                    <div class="bg-white rounded-xl shadow hover:shadow-md transition-shadow flex flex-col h-full border-t-4" style="border-top-color: #094C61;">
                        <div class="p-4 flex-grow flex flex-col">
                            <div class="flex items-center justify-between mb-2">
                                <span class="text-[10px] font-bold px-2 py-0.5 rounded bg-gray-100 text-gray-600">#${index + 1}</span>
                                <i data-lucide="linkedin" class="w-4 h-4 text-blue-600"></i>
                            </div>
                            <div class="custom-scrollbar overflow-y-auto h-24 text-sm text-gray-700 bg-transparent">
                                ${title}
                            </div>
                        </div>
                        <div class="bg-gray-50 p-3 rounded-b-xl border-t border-gray-100 grid grid-cols-2 gap-2">
                            <div>
                                <p class="text-[10px] text-gray-400 uppercase">Impressions</p>
                                <div class="text-sm font-bold text-[#30496D]">${impressions}</div>
                            </div>
                            <div>
                                <p class="text-[10px] text-gray-400 uppercase">Engagement</p>
                                <div class="text-sm font-bold text-[#F46262]">${engRate.toFixed(2)}%</div>
                            </div>
                        </div>
                    </div>
                `;
                postsContainer.insertAdjacentHTML('beforeend', card);
            });
            lucide.createIcons();
        }

        function finalizeDataLoad() {
            // Calculate Start Count
            if (globalSnapshotTotal !== null && globalProcessedData.length > 0) {
                const totalGains = globalProcessedData.reduce((acc, d) => acc + d.val, 0);
                const autoStart = Math.max(0, globalSnapshotTotal - totalGains);
                startCountInput.value = autoStart;
                autoBadge.classList.remove('hidden');
            } else {
                 autoBadge.classList.add('hidden');
            }

            uploadSection.classList.add('hidden');
            dashboard.classList.remove('hidden');
            updateDashboard();
        }

        function updateDashboard() {
            const startCount = parseFloat(startCountInput.value) || 0;
            
            let cumulative = startCount;
            const cumulativeData = [];
            const dailyData = [];
            const labels = [];
            
            let totalGain = 0;

            globalProcessedData.forEach(d => {
                totalGain += d.val;
                cumulative += d.val;
                
                cumulativeData.push(cumulative);
                dailyData.push(d.val);
                
                const dateStr = d.date.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
                labels.push(dateStr);
            });

            // Update KPI
            kpiNewFollowers.textContent = "+" + totalGain.toLocaleString();

            renderChart(labels, dailyData, cumulativeData);
        }

        function renderChart(labels, dailyData, cumulativeData) {
            const ctx = document.getElementById('growthChart').getContext('2d');
            
            if (chartInstance) chartInstance.destroy();

            chartInstance = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: labels,
                    datasets: [
                        {
                            label: 'Daily Gain',
                            data: dailyData,
                            backgroundColor: BRAND_COLORS.munsellBlue + 'B3', 
                            borderColor: BRAND_COLORS.munsellBlue,
                            borderWidth: 1,
                            yAxisID: 'y',
                            borderRadius: 4,
                            order: 2,
                            hoverBackgroundColor: BRAND_COLORS.midnightBlue
                        },
                        {
                            label: 'Total Followers',
                            data: cumulativeData,
                            type: 'line',
                            borderColor: BRAND_COLORS.onixRed, 
                            backgroundColor: BRAND_COLORS.onixRed + '1A',
                            borderWidth: 3,
                            pointRadius: 0,
                            pointHoverRadius: 6,
                            pointHoverBackgroundColor: BRAND_COLORS.onixRed,
                            pointHoverBorderColor: '#fff',
                            fill: true,
                            tension: 0.3,
                            yAxisID: 'y1',
                            order: 1
                        }
                    ]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    interaction: {
                        mode: 'index',
                        intersect: false,
                    },
                    plugins: {
                        legend: {
                            display: false 
                        },
                        tooltip: {
                            backgroundColor: 'rgba(255, 255, 255, 0.95)',
                            titleColor: BRAND_COLORS.satinBlack,
                            bodyColor: '#4b5563',
                            borderColor: '#e5e7eb',
                            borderWidth: 1,
                            padding: 10,
                            displayColors: true,
                        }
                    },
                    scales: {
                        x: {
                            grid: {
                                display: false
                            },
                            ticks: {
                                maxTicksLimit: 12,
                                maxRotation: 0,
                                color: '#6b7280'
                            }
                        },
                        y: {
                            type: 'linear',
                            display: true,
                            position: 'left',
                            title: {
                                display: true,
                                text: 'Daily Gain',
                                color: BRAND_COLORS.munsellBlue
                            },
                            grid: {
                                color: '#f3f4f6'
                            }
                        },
                        y1: {
                            type: 'linear',
                            display: true,
                            position: 'right',
                            title: {
                                display: true,
                                text: 'Total Count',
                                color: BRAND_COLORS.onixRed
                            },
                            grid: {
                                display: false
                            }
                        }
                    }
                }
            });
        }
    </script>
</body>
</html>
