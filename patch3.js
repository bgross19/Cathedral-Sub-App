const fs = require('fs');
let code = fs.readFileSync('Index.html', 'utf8');

// First block in handleAdminAssign
const target1 = `              // Check if At A Glance dashboard is currently visible
              const atAGlanceView = document.getElementById('atAGlanceDashboardView');
              const isAtAGlanceVisible = atAGlanceView && !atAGlanceView.classList.contains('hidden');

              // Refresh admin data behind the scenes
              EventBus.publish('refresh_adminData', { renderAtAGlance: isAtAGlanceVisible });

              // Also refresh quick cover if it's the same 2-day window
              EventBus.publish('refresh_quickCover');`;

const replace1 = `              // Also refresh quick cover if it's the same 2-day window
              EventBus.publish('refresh_quickCover');

              updateLocalAdminData(absenceId, period, subName);`;

code = code.replace(target1, replace1);

// Second block in handleAdminAssign (force block)
const target2 = `                          const atAGlanceVisible = document.getElementById('atAGlanceSection') && !document.getElementById('atAGlanceSection').classList.contains('hidden');
                          const componentToRefresh = atAGlanceVisible ? 'adminData' : 'adminData';

                          google.script.run.withSuccessHandler(refreshed => {
                             if (refreshed.adminData) {
                                adminDataRaw = refreshed.adminData;
                                if (atAGlanceVisible) {
                                   renderAtAGlanceTable(adminDataRaw);
                                } else {
                                   renderAdminDataRawSilent(adminDataRaw);
                                }
                             }
                             updateDashboardTotals();
                          }).refreshData([componentToRefresh]);`;

const replace2 = `                          updateLocalAdminData(absenceId, period, subName);`;

code = code.replace(target2, replace2);

fs.writeFileSync('Index.html', code);
