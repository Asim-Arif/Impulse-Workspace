// Apex Global CRM Chart.js & UI Interop Utilities
window.crmCharts = {
    instances: {},

    waitForChartAndElement: function (canvasId, callback, maxRetries) {
        maxRetries = maxRetries || 30;
        var retries = 0;
        var check = function () {
            var canvas = document.getElementById(canvasId);
            if (typeof Chart !== 'undefined' && canvas) {
                try {
                    callback(canvas);
                } catch (e) {
                    console.error('[crmCharts] Error rendering chart on #' + canvasId + ':', e);
                }
            } else if (retries < maxRetries) {
                retries++;
                setTimeout(check, 80);
            } else {
                console.warn('[crmCharts] Timed out waiting for Chart.js or canvas #' + canvasId + '. Chart=' + (typeof Chart) + ', canvas=' + !!canvas);
            }
        };
        check();
    },

    renderCountryChart: function (canvasId, labels, data) {
        var self = this;
        this.waitForChartAndElement(canvasId, function (canvas) {
            if (self.instances[canvasId]) {
                try { self.instances[canvasId].destroy(); } catch (e) {}
            }

            var colors = [
                '#2563eb', '#3b82f6', '#60a5fa', '#10b981', '#f59e0b',
                '#ef4444', '#8b5cf6', '#ec4899', '#06b6d4', '#64748b'
            ];

            self.instances[canvasId] = new Chart(canvas, {
                type: 'doughnut',
                data: {
                    labels: labels,
                    datasets: [{
                        data: data,
                        backgroundColor: colors.slice(0, labels.length),
                        borderWidth: 2,
                        borderColor: '#ffffff'
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: {
                            position: 'bottom',
                            labels: {
                                boxWidth: 12,
                                font: { size: 11, family: "'Plus Jakarta Sans', sans-serif" }
                            }
                        },
                        tooltip: {
                            callbacks: {
                                label: function (context) {
                                    var val = context.raw || 0;
                                    return ' ' + context.label + ': $' + Number(val).toLocaleString();
                                }
                            }
                        }
                    },
                    cutout: '65%'
                }
            });
        });
    },

    renderPipelineChart: function (canvasId, labels, data) {
        var self = this;
        this.waitForChartAndElement(canvasId, function (canvas) {
            if (self.instances[canvasId]) {
                try { self.instances[canvasId].destroy(); } catch (e) {}
            }

            self.instances[canvasId] = new Chart(canvas, {
                type: 'bar',
                data: {
                    labels: labels,
                    datasets: [{
                        label: 'Deals / Volume',
                        data: data,
                        backgroundColor: [
                            '#64748b', '#3b82f6', '#06b6d4', '#f59e0b', '#8b5cf6', '#10b981'
                        ],
                        borderRadius: 6
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: { display: false },
                        tooltip: {
                            callbacks: {
                                label: function (context) {
                                    return ' Active Deals: ' + context.raw;
                                }
                            }
                        }
                    },
                    scales: {
                        y: {
                            beginAtZero: true,
                            grid: { color: '#f1f5f9' },
                            ticks: { font: { size: 10 } }
                        },
                        x: {
                            grid: { display: false },
                            ticks: { font: { size: 10 } }
                        }
                    }
                }
            });
        });
    },

    renderAgingChart: function (canvasId, labels, data) {
        var self = this;
        this.waitForChartAndElement(canvasId, function (canvas) {
            if (self.instances[canvasId]) {
                try { self.instances[canvasId].destroy(); } catch (e) {}
            }

            self.instances[canvasId] = new Chart(canvas, {
                type: 'bar',
                data: {
                    labels: labels,
                    datasets: [{
                        label: 'Outstanding ($)',
                        data: data,
                        backgroundColor: [
                            '#10b981', '#06b6d4', '#f59e0b', '#f97316', '#ef4444'
                        ],
                        borderRadius: 6
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: { display: false }
                    },
                    scales: {
                        y: {
                            beginAtZero: true,
                            ticks: {
                                callback: function (val) {
                                    return '$' + Number(val / 1000).toLocaleString() + 'k';
                                }
                            }
                        }
                    }
                }
            });
        });
    },

    renderExecutiveRevenueChart: function (canvasId, labels, revenueData, ordersData) {
        var self = this;
        this.waitForChartAndElement(canvasId, function (canvas) {
            if (self.instances[canvasId]) {
                try { self.instances[canvasId].destroy(); } catch (e) {}
            }

            var ctx = canvas.getContext('2d');
            var gradient = ctx.createLinearGradient(0, 0, 0, 300);
            gradient.addColorStop(0, 'rgba(79, 70, 229, 0.45)');
            gradient.addColorStop(0.7, 'rgba(99, 102, 241, 0.1)');
            gradient.addColorStop(1, 'rgba(99, 102, 241, 0.0)');

            self.instances[canvasId] = new Chart(canvas, {
                type: 'line',
                data: {
                    labels: labels,
                    datasets: [
                        {
                            label: 'Gross Sales ($)',
                            data: revenueData,
                            borderColor: '#4f46e5',
                            backgroundColor: gradient,
                            borderWidth: 3,
                            fill: true,
                            tension: 0.35,
                            pointBackgroundColor: '#4f46e5',
                            pointBorderColor: '#ffffff',
                            pointBorderWidth: 2,
                            pointRadius: 4,
                            pointHoverRadius: 6,
                            yAxisID: 'y'
                        },
                        {
                            label: 'Orders Executed',
                            data: ordersData,
                            type: 'bar',
                            backgroundColor: 'rgba(16, 185, 129, 0.75)',
                            borderColor: '#10b981',
                            borderWidth: 1,
                            borderRadius: 6,
                            barThickness: 18,
                            yAxisID: 'y1'
                        }
                    ]
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
                            position: 'top',
                            align: 'end',
                            labels: {
                                boxWidth: 12,
                                font: { size: 11, family: "'Plus Jakarta Sans', sans-serif" },
                                usePointStyle: true
                            }
                        },
                        tooltip: {
                            backgroundColor: 'rgba(15, 23, 42, 0.9)',
                            padding: 10,
                            titleFont: { size: 12, weight: 'bold' },
                            bodyFont: { size: 11 },
                            callbacks: {
                                label: function (context) {
                                    if (context.dataset.yAxisID === 'y') {
                                        return ' Gross Revenue: $' + Number(context.raw).toLocaleString();
                                    }
                                    return ' Orders Executed: ' + context.raw;
                                }
                            }
                        }
                    },
                    scales: {
                        y: {
                            type: 'linear',
                            display: true,
                            position: 'left',
                            beginAtZero: true,
                            grid: { color: 'rgba(226, 232, 240, 0.6)' },
                            ticks: {
                                font: { size: 10 },
                                callback: function (val) {
                                    return '$' + Number(val / 1000).toLocaleString() + 'k';
                                }
                            }
                        },
                        y1: {
                            type: 'linear',
                            display: true,
                            position: 'right',
                            beginAtZero: true,
                            grid: { drawOnChartArea: false },
                            ticks: {
                                font: { size: 10 },
                                stepSize: 5
                            }
                        },
                        x: {
                            grid: { display: false },
                            ticks: { font: { size: 11 } }
                        }
                    }
                }
            });
            console.log('[crmCharts] Successfully initialized #' + canvasId);
        });
    },

    renderExecutiveDonutChart: function (canvasId, labels, data) {
        var self = this;
        this.waitForChartAndElement(canvasId, function (canvas) {
            if (self.instances[canvasId]) {
                try { self.instances[canvasId].destroy(); } catch (e) {}
            }

            var colors = ['#4f46e5', '#06b6d4', '#10b981', '#f59e0b', '#ec4899', '#8b5cf6', '#64748b'];

            self.instances[canvasId] = new Chart(canvas, {
                type: 'doughnut',
                data: {
                    labels: labels,
                    datasets: [{
                        data: data,
                        backgroundColor: colors.slice(0, labels.length),
                        borderWidth: 3,
                        borderColor: '#ffffff',
                        hoverOffset: 6
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: {
                            position: 'bottom',
                            labels: {
                                boxWidth: 10,
                                font: { size: 10, family: "'Plus Jakarta Sans', sans-serif" },
                                padding: 10
                            }
                        },
                        tooltip: {
                            callbacks: {
                                label: function (context) {
                                    var val = context.raw || 0;
                                    return ' ' + context.label + ': $' + Number(val).toLocaleString();
                                }
                            }
                        }
                    },
                    cutout: '68%'
                }
            });
            console.log('[crmCharts] Successfully initialized #' + canvasId);
        });
    },

    renderOfficeTaskChart: function (canvasId, pending, inProgress, completed) {
        var self = this;
        this.waitForChartAndElement(canvasId, function (canvas) {
            if (self.instances[canvasId]) {
                try { self.instances[canvasId].destroy(); } catch (e) {}
            }

            var total = pending + inProgress + completed;
            var displayData = [pending, inProgress, completed];
            if (total === 0) {
                displayData = [1, 1, 1];
            }

            self.instances[canvasId] = new Chart(canvas, {
                type: 'doughnut',
                data: {
                    labels: ['Pending', 'In Progress', 'Completed'],
                    datasets: [{
                        data: displayData,
                        backgroundColor: ['#f59e0b', '#3b82f6', '#10b981'],
                        borderWidth: 2,
                        borderColor: '#ffffff',
                        hoverOffset: 5
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: {
                            display: false
                        },
                        tooltip: {
                            callbacks: {
                                label: function (context) {
                                    return ' ' + context.label + ': ' + context.raw + ' tasks';
                                }
                            }
                        }
                    },
                    cutout: '72%'
                }
            });
            console.log('[crmCharts] Successfully rendered #' + canvasId + ' (P:' + pending + ', IP:' + inProgress + ', C:' + completed + ')');
        });
    }
};
