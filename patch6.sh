cat << 'DIFF' | patch Index.html
--- Index.html
+++ Index.html
@@ -2958,7 +2958,7 @@
               } else {
                  showToast('Sub Removed Successfully', 'success');
               }
-              EventBus.publish('refresh_adminData', { renderAtAGlance: true });
+              updateLocalAdminData(absenceId, period, subName);
             } else {
               showToast(result ? result.error : "Unknown error", 'error');
               btn.disabled = false;
@@ -2979,12 +2979,7 @@
                           } else {
                              showToast('Sub Removed Successfully', 'success');
                           }
-                          google.script.run.withSuccessHandler(refreshed => {
-                             if (refreshed.adminData) {
-                                adminDataRaw = refreshed.adminData;
-                                renderAtAGlanceTable(adminDataRaw);
-                             }
-                          }).refreshData(['adminData']);
+                          updateLocalAdminData(absenceId, period, subName);
                        }
                     })
                     .withFailureHandler(forceErr => {
DIFF
