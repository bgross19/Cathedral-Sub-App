cat << 'DIFF' | patch Index.html
--- Index.html
+++ Index.html
@@ -4211,8 +4211,8 @@
               if (isFromModal) {
                  closeDetailsModal();
               }
-              EventBus.publish('refresh_quickCover'); // Refresh the list
-              EventBus.publish('refresh_adminData');
+              EventBus.publish('refresh_quickCover');
+              updateLocalAdminData(absenceId, period, subName);
             } else {
               showToast(result ? result.error : "Unknown error", 'error');
               btn.disabled = false;
@@ -4239,14 +4239,8 @@
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
