cat << 'DIFF' | patch Index.html
--- Index.html
+++ Index.html
@@ -2905,14 +2905,8 @@
                           if (isFromModal) {
                              closeAllModals();
                           }
-                          google.script.run.withSuccessHandler(refreshed => {
-                             if (refreshed.quickCover) {
-                                renderQuickCover(refreshed.quickCover);
-                             }
-                             if (refreshed.adminData) {
-                                adminDataRaw = refreshed.adminData;
-                             }
-                          }).refreshData(['quickCover', 'adminData']);
+                          EventBus.publish('refresh_quickCover');
+                          updateLocalAdminData(absenceId, period, subName);
                        }
                     })
                     .withFailureHandler(forceErr => {
DIFF
