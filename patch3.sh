cat << 'DIFF' | patch Index.html
--- Index.html
+++ Index.html
@@ -3955,11 +3955,8 @@
               } else {
                  showToast('Sub Removed Successfully', 'success');
               }
-
-              // Check if At A Glance dashboard is currently visible
-              const atAGlanceView = document.getElementById('atAGlanceDashboardView');
-              const isAtAGlanceVisible = atAGlanceView && !atAGlanceView.classList.contains('hidden');
-
-              // Refresh admin data behind the scenes
-              EventBus.publish('refresh_adminData', { renderAtAGlance: isAtAGlanceVisible });
-
               // Also refresh quick cover if it's the same 2-day window
               EventBus.publish('refresh_quickCover');
+
+              updateLocalAdminData(absenceId, period, subName);
             } else {
               showToast(result ? result.error : "Unknown error", 'error');
               btn.disabled = false;
@@ -3990,16 +3987,7 @@
                           } else {
                              showToast('Sub Removed Successfully', 'success');
                           }
-                          const atAGlanceVisible = document.getElementById('atAGlanceSection') && !document.getElementById('atAGlanceSection').classList.contains('hidden');
-                          const componentToRefresh = atAGlanceVisible ? 'adminData' : 'adminData';
-
-                          google.script.run.withSuccessHandler(refreshed => {
-                             if (refreshed.adminData) {
-                                adminDataRaw = refreshed.adminData;
-                                if (atAGlanceVisible) {
-                                   renderAtAGlanceTable(adminDataRaw);
-                                } else {
-                                   renderAdminDataRawSilent(adminDataRaw);
-                                }
-                             }
-                             updateDashboardTotals();
-                          }).refreshData([componentToRefresh]);
+
+                          updateLocalAdminData(absenceId, period, subName);
                        } else {
                           showToast('Failed to assign sub.', 'error');
                        }
DIFF
