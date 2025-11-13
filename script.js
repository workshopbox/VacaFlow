let employeesData = [];
        let filteredData = [];
        let weekDefinitions = [];
        let allCharts = {};
        
        // Define all day types based on Excel classes
        const DAY_TYPES = {
            vacation: {
                class: 'xl68',
                color: 'lime',
                label: 'Vacation',
                icon: '🏖️'
            },
            flextime: {
                class: 'xl70',
                color: '#7F00BF',
                label: 'Flextime Day',
                icon: '⏰'
            },
            papamonat: {
                class: 'xl71',
                color: '#FF3FFF',
                label: 'Papamonat',
                icon: '👶'
            },
            specialLeave: {
                class: 'xl72',
                color: '#00FFBF',
                label: 'Special Leave',
                icon: '📋'
            },
            longSick: {
                class: 'xl73',
                color: 'red',
                label: 'Long Sick Leave',
                icon: '🏥'
            },
            shortSick: {
                class: 'xl74',
                color: '#BF7F00',
                label: 'Short Sick Leave',
                icon: '🤒'
            },
            activeTravel: {
                class: 'xl75',
                color: '#7FFFFF',
                label: 'Active Travel Time',
                icon: '✈️'
            }
        };
        
        const fileInput = document.getElementById('fileInput');
        const fileInfo = document.getElementById('fileInfo');
        const processInfo = document.getElementById('processInfo');
        const uploadSection = document.getElementById('uploadSection');

        // Drag and drop functionality
        uploadSection.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadSection.classList.add('dragover');
        });

        uploadSection.addEventListener('dragleave', () => {
            uploadSection.classList.remove('dragover');
        });

        uploadSection.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadSection.classList.remove('dragover');
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                fileInput.files = files;
                handleFile(files[0]);
            }
        });

        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                handleFile(e.target.files[0]);
            }
        });

        function handleFile(file) {
            if (!file.name.match(/\.(htm|html)$/i)) {
                alert('Please upload an HTML file (.htm or .html)');
                return;
            }

            fileInfo.textContent = `Processing: ${file.name}...`;
            fileInfo.classList.add('show');
            processInfo.classList.add('show');
            document.getElementById('loadingOverlay').classList.add('show');
            document.getElementById('loadingText').textContent = 'Reading HTML file...';
            document.getElementById('loadingDetail').textContent = 'Extracting colored cells';

            // Read HTML file as text
            const reader = new FileReader();
            reader.onload = async function(e) {
                try {
                    const htmlText = e.target.result;
                    await processHTMLFile(htmlText, file.name);
                } catch (error) {
                    alert('Error reading file: ' + error.message);
                    console.error(error);
                    document.getElementById('loadingOverlay').classList.remove('show');
                }
            };
            reader.readAsText(file);
        }

        async function processHTMLFile(htmlText, filename) {
            try {
                // Parse HTML
                const parser = new DOMParser();
                const doc = parser.parseFromString(htmlText, 'text/html');
                const table = doc.querySelector('table');
                
                if (!table) {
                    throw new Error('No table found in HTML file');
                }

                const rows = Array.from(table.querySelectorAll('tr'));
                
                // Find header row with week information
                const headerRow = rows[0];
                const weekCells = Array.from(headerRow.querySelectorAll('td')).filter(cell => 
                    cell.textContent.toLowerCase().includes('week')
                );
                
                document.getElementById('loadingText').textContent = 'Detecting colored cells...';
                
                // Parse week definitions
                weekDefinitions = [];
                let currentCol = 0;
                
                for (const row of rows) {
                    const cells = Array.from(row.querySelectorAll('td'));
                    if (cells.length > 0 && cells[0].textContent.toLowerCase().includes('week')) {
                        let colIndex = 2; // Start after Department and Name columns
                        for (let i = 0; i < cells.length; i++) {
                            const cellText = cells[i].textContent.trim();
                            if (cellText.match(/Week \d+/i)) {
                                const weekNum = parseInt(cellText.match(/\d+/)[0]);
                                const colspan = parseInt(cells[i].getAttribute('colspan') || '7');
                                weekDefinitions.push({
                                    week: weekNum,
                                    label: cellText,
                                    startCol: colIndex,
                                    endCol: colIndex + colspan - 1
                                });
                                colIndex += colspan;
                            }
                        }
                        break;
                    }
                }

                // If no weeks found, create default structure
                if (weekDefinitions.length === 0) {
                    const dayColumns = rows[1]?.querySelectorAll('td').length - 2 || 30;
                    const daysPerWeek = Math.ceil(dayColumns / 4);
                    for (let i = 0; i < 4; i++) {
                        weekDefinitions.push({
                            week: i + 1,
                            label: `Week ${i + 1}`,
                            startCol: 2 + (i * daysPerWeek),
                            endCol: 2 + ((i + 1) * daysPerWeek) - 1
                        });
                    }
                }

                document.getElementById('loadingText').textContent = 'Processing employees...';
                
                // Process employee data
                employeesData = [];
                const dataRows = rows.slice(2); // Skip header rows
                
                for (const row of dataRows) {
                    const cells = Array.from(row.querySelectorAll('td'));
                    if (cells.length < 3) continue;
                    
                    const deptText = cells[0]?.textContent.trim() || '';
                    const nameText = cells[1]?.textContent.trim() || '';
                    
                    if (!nameText || nameText.toLowerCase() === 'name') continue;
                    
                    // Parse department column: "DAP8 JL1 Früh (DAP8 JL1 Früh)" or "Job Level 4 (JL4 DAP8)"
                    let department = 'Unknown';
                    let jobLevel = 'Unknown';
                    let shiftFromDept = '';
                    
                    // Extract from parentheses first (more reliable)
                    const parenMatch = deptText.match(/\(([^)]+)\)/);
                    const deptInfo = parenMatch ? parenMatch[1].trim() : deptText.split('\n')[0].trim();
                    
                    // Extract department (first 4 chars, e.g., DAP8)
                    const deptMatch = deptInfo.match(/([A-Z]{3}\d+)/);
                    if (deptMatch) {
                        department = deptMatch[1];
                    }
                    
                    // Extract job level (JL1, JL3, JL4, JL5, etc.)
                    const jlMatch = deptInfo.match(/JL\s*(\d+)/i) || deptText.match(/Job Level\s*(\d+)/i);
                    if (jlMatch) {
                        jobLevel = 'JL' + jlMatch[1];
                    }
                    
                    // Extract shift - look at the NAME column which has "Frühschicht", "Nachtschicht", etc.
                    let shift = 'Unknown';
                    if (nameText.match(/schicht\s*\d/i)) {
                        // Check what type of schicht
                        if (nameText.toLowerCase().includes('nacht')) {
                            shift = 'Nacht';
                        } else if (nameText.toLowerCase().includes('hybrid')) {
                            shift = 'Hybrid';
                        } else if (nameText.toLowerCase().includes('sp')) {
                            shift = 'Spät';
                        } else {
                            // If it contains "schicht" but not the above, it's likely Früh
                            shift = 'Früh';
                        }
                    } else if (nameText.match(/ORM/i) || deptInfo.match(/ORM/i)) {
                        shift = 'ORM';
                    } else if (nameText.match(/^L[34567]/)) {
                        // For L3, L4, L5, etc. employees, check department column for shift
                        if (deptInfo.toLowerCase().includes('nacht')) {
                            shift = 'Nacht';
                        } else if (deptInfo.toLowerCase().includes('hybrid')) {
                            shift = 'Hybrid';
                        } else if (deptInfo.toLowerCase().includes('sp')) {
                            shift = 'Spät';
                        } else if (deptInfo.toLowerCase().includes('orm')) {
                            shift = 'ORM';
                        } else {
                            // Default for managers/leads is Früh
                            shift = 'Früh';
                        }
                    }
                    
                    // Clean up name - remove parentheses content
                    let name = nameText.replace(/\([^)]+\)/, '').trim();
                    
                    // Parse day cells for all day types
                    const days = [];
                    for (let i = 2; i < cells.length; i++) {
                        const cell = cells[i];
                        const cellClass = cell.getAttribute('class') || '';
                        const textContent = cell.textContent.trim();
                        
                        // Detect which day type this is
                        let dayType = null;
                        for (const [key, config] of Object.entries(DAY_TYPES)) {
                            if (cellClass.includes(config.class)) {
                                dayType = key;
                                break;
                            }
                        }
                        
                        days.push({
                            index: i - 2,
                            class: cellClass,
                            dayType: dayType,
                            text: textContent
                        });
                    }
                    
                    employeesData.push({
                        department: department,
                        name: name,
                        shift: shift,
                        job_level: jobLevel,
                        days: days
                    });
                }

                if (employeesData.length === 0) {
                    throw new Error('No employee data found in file');
                }

                filteredData = [...employeesData];
                
                // Update UI
                populateFilters();
                updateDashboard();
                
                document.getElementById('filtersSection').classList.add('show');
                document.getElementById('dashboard').classList.add('show');
                
                fileInfo.textContent = `✓ Successfully loaded ${employeesData.length} employees`;
                processInfo.textContent = `📊 Found ${weekDefinitions.length} weeks of data`;
                
                document.getElementById('loadingOverlay').classList.remove('show');
                
            } catch (error) {
                console.error('Error processing HTML:', error);
                alert('Error processing file: ' + error.message);
                document.getElementById('loadingOverlay').classList.remove('show');
            }
        }

        function countDaysByType(employee, dayType, weekNum = null) {
            let count = 0;
            for (const day of employee.days) {
                if (day.dayType === dayType) {
                    if (weekNum !== null) {
                        // Check if day is in the specified week
                        const week = weekDefinitions.find(w => w.week === weekNum);
                        if (week && day.index >= week.startCol - 2 && day.index <= week.endCol - 2) {
                            count++;
                        }
                    } else {
                        count++;
                    }
                }
            }
            return count;
        }
        
        function countVacationDays(employee, weekNum = null) {
            return countDaysByType(employee, 'vacation', weekNum);
        }
        
        function countTotalAbsenceDays(employee, weekNum = null) {
            let count = 0;
            for (const type of Object.keys(DAY_TYPES)) {
                count += countDaysByType(employee, type, weekNum);
            }
            return count;
        }

        function populateFilters() {
            const departments = [...new Set(employeesData.map(e => e.department))].sort();
            const shifts = [...new Set(employeesData.map(e => e.shift))].sort();
            const levels = [...new Set(employeesData.map(e => e.job_level))].sort();
            
            const deptFilter = document.getElementById('deptFilter');
            const shiftFilter = document.getElementById('shiftFilter');
            const levelFilter = document.getElementById('levelFilter');
            
            deptFilter.innerHTML = '<option value="all">All Departments</option>';
            departments.forEach(dept => {
                deptFilter.innerHTML += `<option value="${dept}">${dept}</option>`;
            });
            
            shiftFilter.innerHTML = '<option value="all">All Shifts</option>';
            shifts.forEach(shift => {
                shiftFilter.innerHTML += `<option value="${shift}">${shift}</option>`;
            });
            
            levelFilter.innerHTML = '<option value="all">All Levels</option>';
            levels.forEach(level => {
                levelFilter.innerHTML += `<option value="${level}">${level}</option>`;
            });
        }

        function updateDashboard() {
            const stats = calculateStats();
            updateStatsCards(stats);
            createCharts(stats);
            updateTable();
        }

        function calculateStats() {
            const totalEmployees = filteredData.length;
            
            // Calculate totals for each day type
            const dayTypeTotals = {};
            for (const [key, config] of Object.entries(DAY_TYPES)) {
                dayTypeTotals[key] = {
                    total: filteredData.reduce((sum, emp) => sum + countDaysByType(emp, key), 0),
                    label: config.label,
                    color: config.color,
                    icon: config.icon
                };
            }
            
            const totalVacationDays = dayTypeTotals.vacation.total;
            const totalAbsenceDays = filteredData.reduce((sum, emp) => sum + countTotalAbsenceDays(emp), 0);
            const avgVacationPerEmployee = totalEmployees > 0 ? (totalVacationDays / totalEmployees).toFixed(1) : 0;
            
            // Calculate vacation quote (percentage of days that are vacations)
            const totalPossibleDays = totalEmployees * (filteredData[0]?.days.length || 30);
            const vacationQuote = totalPossibleDays > 0 ? ((totalVacationDays / totalPossibleDays) * 100).toFixed(1) : 0;
            
            // Find employee with most vacations
            let empWithMostVacations = { name: '', days: 0 };
            filteredData.forEach(emp => {
                const days = countVacationDays(emp);
                if (days > empWithMostVacations.days) {
                    empWithMostVacations = { name: emp.name, days };
                }
            });
            
            // Weekly data for ALL day types
            const weeklyData = weekDefinitions.map(week => {
                const weekStats = { label: week.label };
                for (const [key, config] of Object.entries(DAY_TYPES)) {
                    weekStats[key] = filteredData.reduce((sum, emp) => sum + countDaysByType(emp, key, week.week), 0);
                }
                return weekStats;
            });
            
            return {
                totalEmployees,
                totalVacationDays,
                totalAbsenceDays,
                avgVacationPerEmployee,
                vacationQuote,
                empWithMostVacations,
                weeklyData,
                dayTypeTotals
            };
        }

        function updateStatsCards(stats) {
            const currentMonth = new Date().toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
            
            let cardsHTML = `
                <div class="stat-card">
                    <div class="stat-header">
                        <div class="stat-label">Total Employees</div>
                        <div class="stat-icon">👥</div>
                    </div>
                    <div class="stat-value">${stats.totalEmployees}</div>
                    <div class="stat-sublabel">Active employees</div>
                </div>
                <div class="stat-card">
                    <div class="stat-header">
                        <div class="stat-label">Total Absence Days</div>
                        <div class="stat-icon">📊</div>
                    </div>
                    <div class="stat-value">${stats.totalAbsenceDays}</div>
                    <div class="stat-sublabel">All types combined</div>
                </div>
            `;
            
            // Add a card for each day type
            for (const [key, data] of Object.entries(stats.dayTypeTotals)) {
                if (data.total > 0) {
                    cardsHTML += `
                    <div class="stat-card">
                        <div class="stat-header">
                            <div class="stat-label">${data.label}</div>
                            <div class="stat-icon">${data.icon}</div>
                        </div>
                        <div class="stat-value">${data.total}</div>
                        <div class="stat-sublabel">${currentMonth}</div>
                    </div>
                    `;
                }
            }
            
            document.getElementById('statsGrid').innerHTML = cardsHTML;
        }

        function createCharts(stats) {
            // Destroy existing charts
            Object.values(allCharts).forEach(chart => chart?.destroy());
            
            // Chart 1: Day Types Distribution (Pie/Doughnut)
            const dayTypeLabels = [];
            const dayTypeData = [];
            const dayTypeColors = [];
            
            for (const [key, data] of Object.entries(stats.dayTypeTotals)) {
                if (data.total > 0) {
                    dayTypeLabels.push(data.label);
                    dayTypeData.push(data.total);
                    dayTypeColors.push(data.color);
                }
            }
            
            allCharts.dayTypeChart = new Chart(document.getElementById('shiftChart'), {
                type: 'doughnut',
                data: {
                    labels: dayTypeLabels,
                    datasets: [{
                        data: dayTypeData,
                        backgroundColor: dayTypeColors,
                        borderWidth: 2,
                        borderColor: '#fff'
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: { 
                            position: 'bottom',
                            labels: {
                                padding: 10,
                                font: { size: 11 }
                            }
                        },
                        title: {
                            display: true,
                            text: 'Absence Days by Type',
                            font: { size: 14, weight: 'bold' }
                        }
                    }
                }
            });
            
            // Chart 2: Day Types by Shift (Stacked Bar)
            const shifts = [...new Set(filteredData.map(e => e.shift))];
            const datasets = [];
            
            for (const [key, config] of Object.entries(DAY_TYPES)) {
                const shiftData = shifts.map(shift => {
                    return filteredData
                        .filter(emp => emp.shift === shift)
                        .reduce((sum, emp) => sum + countDaysByType(emp, key), 0);
                });
                
                const total = shiftData.reduce((a, b) => a + b, 0);
                if (total > 0) {
                    datasets.push({
                        label: config.label,
                        data: shiftData,
                        backgroundColor: config.color,
                        borderWidth: 1,
                        borderColor: '#fff'
                    });
                }
            }
            
            allCharts.shiftChart = new Chart(document.getElementById('levelChart'), {
                type: 'bar',
                data: {
                    labels: shifts,
                    datasets: datasets
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: { 
                            position: 'bottom',
                            labels: {
                                padding: 8,
                                font: { size: 10 }
                            }
                        },
                        title: {
                            display: true,
                            text: 'Absence Days by Shift',
                            font: { size: 14, weight: 'bold' }
                        }
                    },
                    scales: {
                        x: { stacked: true },
                        y: { 
                            stacked: true,
                            beginAtZero: true
                        }
                    }
                }
            });
            
            // Chart 3: Daily Coverage for ALL types (Stacked Area)
            const maxDays = filteredData.length > 0 ? filteredData[0].days.length : 30;
            const dailyDatasets = [];
            
            for (const [key, config] of Object.entries(DAY_TYPES)) {
                const dailyData = Array(maxDays).fill(0);
                filteredData.forEach(emp => {
                    emp.days.forEach((day, idx) => {
                        if (day.dayType === key && idx < maxDays) {
                            dailyData[idx]++;
                        }
                    });
                });
                
                const total = dailyData.reduce((a, b) => a + b, 0);
                if (total > 0) {
                    dailyDatasets.push({
                        label: config.label,
                        data: dailyData,
                        borderColor: config.color,
                        backgroundColor: config.color + '40', // Add transparency
                        fill: true,
                        tension: 0.4,
                        borderWidth: 2
                    });
                }
            }
            
            allCharts.dailyChart = new Chart(document.getElementById('dailyChart'), {
                type: 'line',
                data: {
                    labels: Array.from({length: maxDays}, (_, i) => `Day ${i + 1}`),
                    datasets: dailyDatasets
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    interaction: {
                        mode: 'index',
                        intersect: false
                    },
                    plugins: {
                        legend: { 
                            position: 'bottom',
                            labels: {
                                padding: 8,
                                font: { size: 10 }
                            }
                        },
                        title: {
                            display: true,
                            text: 'Daily Absence Coverage',
                            font: { size: 14, weight: 'bold' }
                        }
                    },
                    scales: {
                        y: { 
                            stacked: true,
                            beginAtZero: true
                        }
                    }
                }
            });
            
            // Chart 4: Weekly Distribution (Stacked Bar)
            const weeklyDatasets = [];
            
            for (const [key, config] of Object.entries(DAY_TYPES)) {
                const weeklyData = stats.weeklyData.map(week => week[key] || 0);
                const total = weeklyData.reduce((a, b) => a + b, 0);
                
                if (total > 0) {
                    weeklyDatasets.push({
                        label: config.label,
                        data: weeklyData,
                        backgroundColor: config.color,
                        borderWidth: 1,
                        borderColor: '#fff'
                    });
                }
            }
            
            allCharts.weeklyChart = new Chart(document.getElementById('weeklyChart'), {
                type: 'bar',
                data: {
                    labels: stats.weeklyData.map(w => w.label),
                    datasets: weeklyDatasets
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: { 
                            position: 'bottom',
                            labels: {
                                padding: 8,
                                font: { size: 10 }
                            }
                        },
                        title: {
                            display: true,
                            text: 'Weekly Absence Distribution',
                            font: { size: 14, weight: 'bold' }
                        }
                    },
                    scales: {
                        x: { stacked: true },
                        y: { 
                            stacked: true,
                            beginAtZero: true
                        }
                    }
                }
            });
        }

        function updateTable() {
            const tbody = document.getElementById('tableBody');
            tbody.innerHTML = '';
            
            filteredData.forEach(emp => {
                const row = tbody.insertRow();
                const shiftClass = `shift-${emp.shift.toLowerCase()}`;
                
                // Count each day type
                const counts = {};
                for (const key of Object.keys(DAY_TYPES)) {
                    counts[key] = countDaysByType(emp, key);
                }
                const total = countTotalAbsenceDays(emp);
                
                row.innerHTML = `
                    <td><strong>${emp.name}</strong></td>
                    <td><span class="shift-badge ${shiftClass}">${emp.shift}</span></td>
                    <td><span class="level-badge">${emp.job_level}</span></td>
                    <td style="background-color: ${DAY_TYPES.vacation.color}20;">${counts.vacation || '-'}</td>
                    <td style="background-color: ${DAY_TYPES.flextime.color}20;">${counts.flextime || '-'}</td>
                    <td style="background-color: ${DAY_TYPES.papamonat.color}20;">${counts.papamonat || '-'}</td>
                    <td style="background-color: ${DAY_TYPES.specialLeave.color}20;">${counts.specialLeave || '-'}</td>
                    <td style="background-color: ${DAY_TYPES.longSick.color}20;">${counts.longSick || '-'}</td>
                    <td style="background-color: ${DAY_TYPES.shortSick.color}20;">${counts.shortSick || '-'}</td>
                    <td style="background-color: ${DAY_TYPES.activeTravel.color}20;">${counts.activeTravel || '-'}</td>
                    <td><strong>${total}</strong></td>
                `;
            });
        }

        function applyFilters() {
            const deptValue = document.getElementById('deptFilter').value;
            const shiftValue = document.getElementById('shiftFilter').value;
            const levelValue = document.getElementById('levelFilter').value;
            
            filteredData = employeesData.filter(emp => {
                if (deptValue !== 'all' && emp.department !== deptValue) return false;
                if (shiftValue !== 'all' && emp.shift !== shiftValue) return false;
                if (levelValue !== 'all' && emp.job_level !== levelValue) return false;
                return true;
            });
            
            updateDashboard();
        }

        function filterTable() {
            const searchTerm = document.getElementById('searchBox').value.toLowerCase();
            const rows = document.querySelectorAll('#dataTable tbody tr');
            
            rows.forEach(row => {
                const text = row.textContent.toLowerCase();
                row.style.display = text.includes(searchTerm) ? '' : 'none';
            });
        }

        function exportToCSV() {
            let csv = 'Name,Shift,Job Level,Total Vacation Days';
            weekDefinitions.forEach(w => csv += `,${w.label}`);
            csv += '\n';
            
            filteredData.forEach(emp => {
                const weekVacations = weekDefinitions.map(w => countVacationDays(emp, w.week));
                csv += `${emp.name},${emp.shift},${emp.job_level},${countVacationDays(emp)},${weekVacations.join(',')}\n`;
            });
            
            const blob = new Blob([csv], { type: 'text/csv' });
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `vacation_analytics_${new Date().toISOString().split('T')[0]}.csv`;
            a.click();
        }
    