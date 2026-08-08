cat << 'DIFF' | patch Index.html
--- Index.html
+++ Index.html
@@ -4065,11 +4065,8 @@
               inputEl.disabled = false;
               inputEl.classList.remove('bg-[#f3f4f6]');

-              const atAGlanceView = document.getElementById('atAGlanceDashboardView');
-              const isAtAGlanceVisible = atAGlanceView && !atAGlanceView.classList.contains('hidden');
-
-              EventBus.publish('refresh_adminData', { renderAtAGlance: isAtAGlanceVisible });
               EventBus.publish('refresh_quickCover');
+              updateLocalAdminData(absenceId, period, subName);
             } else {
               showToast(result ? result.error : "Unknown error", 'error');
               inputEl.value = originalValue;
@@ -4107,13 +4104,7 @@
                              inputEl.classList.remove('text-green-700', 'font-bold');
                              inputEl.classList.add('text-gray-400', 'italic');
                           }
-
-                          google.script.run.withSuccessHandler(refreshed => {
-                             if (refreshed.adminData) {
-                                adminDataRaw = refreshed.adminData;
-                                updateDashboardTotals();
-                             }
-                          }).refreshData(['adminData']);
+                          updateLocalAdminData(absenceId, period, subName);
                        }
                     })
                     .withFailureHandler(forceErr => {
DIFF
