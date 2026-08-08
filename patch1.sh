cat << 'DIFF' | patch Index.html
--- Index.html
+++ Index.html
@@ -3553,6 +3553,30 @@
           .refreshData(['adminData']);
       }

+      function updateLocalAdminData(absenceId, period, subName) {
+          // Update raw data
+          const idx = adminDataRaw.findIndex(r => String(r.id) === String(absenceId) && String(r.period) === String(period));
+          if (idx !== -1) {
+              adminDataRaw[idx].assignedSub = subName || "";
+          }
+
+          // Also update filtered data if it exists
+          const filterIdx = adminDataFiltered.findIndex(r => String(r.id) === String(absenceId) && String(r.period) === String(period));
+          if (filterIdx !== -1) {
+              adminDataFiltered[filterIdx].assignedSub = subName || "";
+          }
+
+          const atAGlanceView = document.getElementById('atAGlanceDashboardView');
+          const isAtAGlanceVisible = atAGlanceView && !atAGlanceView.classList.contains('hidden');
+
+          if (isAtAGlanceVisible) {
+             applyAtAGlanceFilter();
+          } else {
+             // Full re-render of admin table to reflect new data
+             renderAdminTable();
+          }
+      }
+
       function renderAdminDashboardData(data, renderAtAGlance = false) {
             adminDataRaw = data;
             const today = new Date();
DIFF
