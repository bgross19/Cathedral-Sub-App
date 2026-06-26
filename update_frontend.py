import re

def main():
    with open('Index.html', 'r') as f:
        content = f.read()

    # The functions in Index.html to modify:
    # loadStaffList -> refreshData(['staffList'])
    # loadTodaysOpenJobs -> refreshData(['todaysOpenJobs'])
    # loadQuickCover -> refreshData(['quickCover'])
    # loadMySubDuties -> refreshData(['mySubDuties'])
    # loadMyAbsences -> refreshData(['myAbsences'])

    # loadAdminDataRawSilent -> refreshData(['adminData']) -> calls renderAdminDataRawSilent
    # loadAdminDashboardData -> refreshData(['adminData']) -> calls renderAdminDashboardData
    # loadHRDashboardData -> refreshData(['hrData']) -> calls renderHRDashboardData

    # Let's write a replacement script for each

    # 1. loadStaffList
    target = """      function loadStaffList() {
        google.script.run
          .withSuccessHandler(renderStaffList)
          .getStaffList();
      }"""
    replacement = """      function loadStaffList() {
        google.script.run
          .withSuccessHandler(res => { if (res.staffList) renderStaffList(res.staffList); })
          .refreshData(['staffList']);
      }"""
    content = content.replace(target, replacement)

    # 2. loadTodaysOpenJobs
    target = """      function loadTodaysOpenJobs() {
        google.script.run
          .withSuccessHandler(renderTodaysOpenJobs)
          .withFailureHandler(err => {
            document.getElementById('todaysOpenJobs').innerHTML = `<div class="text-red-600 font-bold">Failed to load open jobs: ${err.message}</div>`;
          })
          .getTodaysOpenJobsData();
      }"""
    replacement = """      function loadTodaysOpenJobs() {
        google.script.run
          .withSuccessHandler(res => { if (res.todaysOpenJobs) renderTodaysOpenJobs(res.todaysOpenJobs); })
          .withFailureHandler(err => {
            document.getElementById('todaysOpenJobs').innerHTML = `<div class="text-red-600 font-bold">Failed to load open jobs: ${err.message}</div>`;
          })
          .refreshData(['todaysOpenJobs']);
      }"""
    content = content.replace(target, replacement)

    # 3. loadQuickCover
    target = """      function loadQuickCover() {
        google.script.run
          .withSuccessHandler(renderQuickCover)
          .withFailureHandler(err => {
            document.getElementById('quickCoverContainer').innerHTML = `<div class="p-8 text-center text-red-600 font-bold">Failed to load data: ${err.message}</div>`;
          })
          .getQuickCoverData();
      }"""
    replacement = """      function loadQuickCover() {
        google.script.run
          .withSuccessHandler(res => { if (res.quickCover) renderQuickCover(res.quickCover); })
          .withFailureHandler(err => {
            document.getElementById('quickCoverContainer').innerHTML = `<div class="p-8 text-center text-red-600 font-bold">Failed to load data: ${err.message}</div>`;
          })
          .refreshData(['quickCover']);
      }"""
    content = content.replace(target, replacement)

    # 4. loadMySubDuties
    target = """      function loadMySubDuties() {
        google.script.run
          .withSuccessHandler(renderMySubDuties)
          .withFailureHandler(err => {
            document.getElementById('mySubDuties').innerHTML = `<div class="p-8 text-center text-red-600 font-bold">Failed to load sub-duties: ${err.message}</div>`;
          })
          .getMySubDuties();
      }"""
    replacement = """      function loadMySubDuties() {
        google.script.run
          .withSuccessHandler(res => { if (res.mySubDuties) renderMySubDuties(res.mySubDuties); })
          .withFailureHandler(err => {
            document.getElementById('mySubDuties').innerHTML = `<div class="p-8 text-center text-red-600 font-bold">Failed to load sub-duties: ${err.message}</div>`;
          })
          .refreshData(['mySubDuties']);
      }"""
    content = content.replace(target, replacement)

    # 5. loadMyAbsences
    target = """      function loadMyAbsences() {
        google.script.run
          .withSuccessHandler(renderMyAbsences)
          .withFailureHandler(err => {
            document.getElementById('myAbsences').innerHTML = `<div class="text-red-600 font-bold">Failed to load absences: ${err.message}</div>`;
          })
          .getMyAbsences();
      }"""
    replacement = """      function loadMyAbsences() {
        google.script.run
          .withSuccessHandler(res => { if (res.myAbsences) renderMyAbsences(res.myAbsences); })
          .withFailureHandler(err => {
            document.getElementById('myAbsences').innerHTML = `<div class="text-red-600 font-bold">Failed to load absences: ${err.message}</div>`;
          })
          .refreshData(['myAbsences']);
      }"""
    content = content.replace(target, replacement)

    # 6. loadAdminDataRawSilent
    target = """      function loadAdminDataRawSilent() {
        google.script.run
          .withSuccessHandler(renderAdminDataRawSilent)
          .getAdminDashboardData();
      }"""
    replacement = """      function loadAdminDataRawSilent() {
        google.script.run
          .withSuccessHandler(res => { if (res.adminData) renderAdminDataRawSilent(res.adminData); })
          .refreshData(['adminData']);
      }"""
    content = content.replace(target, replacement)

    # 7. loadAdminDashboardData
    target = """      function loadAdminDashboardData(renderAtAGlance = false) {
        document.getElementById('adminDashboardTableBody').innerHTML = '<tr><td colspan="7" class="p-8 text-center text-gray-400 italic">Loading dashboard data...</td></tr>';
        if (renderAtAGlance) {
            document.getElementById('atAGlanceTableBody').innerHTML = '<tr><td class="p-8 text-center text-gray-400 italic">Loading Today at a Glance...</td></tr>';
        }

        google.script.run
          .withSuccessHandler(data => renderAdminDashboardData(data, renderAtAGlance))
          .withFailureHandler(err => {
            document.getElementById('adminDashboardTableBody').innerHTML = `<tr><td colspan="7" class="p-8 text-center text-red-600 font-bold">Failed to load: ${err.message}</td></tr>`;
            if (renderAtAGlance) {
                document.getElementById('atAGlanceTableBody').innerHTML = `<tr><td class="p-8 text-center text-red-600 font-bold">Failed to load: ${err.message}</td></tr>`;
            }
          })
          .getAdminDashboardData();
      }"""
    replacement = """      function loadAdminDashboardData(renderAtAGlance = false) {
        document.getElementById('adminDashboardTableBody').innerHTML = '<tr><td colspan="7" class="p-8 text-center text-gray-400 italic">Loading dashboard data...</td></tr>';
        if (renderAtAGlance) {
            document.getElementById('atAGlanceTableBody').innerHTML = '<tr><td class="p-8 text-center text-gray-400 italic">Loading Today at a Glance...</td></tr>';
        }

        google.script.run
          .withSuccessHandler(res => { if (res.adminData) renderAdminDashboardData(res.adminData, renderAtAGlance); })
          .withFailureHandler(err => {
            document.getElementById('adminDashboardTableBody').innerHTML = `<tr><td colspan="7" class="p-8 text-center text-red-600 font-bold">Failed to load: ${err.message}</td></tr>`;
            if (renderAtAGlance) {
                document.getElementById('atAGlanceTableBody').innerHTML = `<tr><td class="p-8 text-center text-red-600 font-bold">Failed to load: ${err.message}</td></tr>`;
            }
          })
          .refreshData(['adminData']);
      }"""
    content = content.replace(target, replacement)

    # 8. loadHRDashboardData
    target = """      function loadHRDashboardData() {
        document.getElementById('hrDashboardTableBody').innerHTML = '<tr><td colspan="2" class="p-8 text-center text-gray-400 italic">Loading HR dashboard data...</td></tr>';

        google.script.run
          .withSuccessHandler(renderHRDashboardData)
          .withFailureHandler(err => {
            document.getElementById('hrDashboardTableBody').innerHTML = `<tr><td colspan="2" class="p-8 text-center text-red-600 font-bold">Failed to load: ${err.message}</td></tr>`;
          })
          .getHRDashboardData();
      }"""
    replacement = """      function loadHRDashboardData() {
        document.getElementById('hrDashboardTableBody').innerHTML = '<tr><td colspan="2" class="p-8 text-center text-gray-400 italic">Loading HR dashboard data...</td></tr>';

        google.script.run
          .withSuccessHandler(res => { if (res.hrData) renderHRDashboardData(res.hrData); })
          .withFailureHandler(err => {
            document.getElementById('hrDashboardTableBody').innerHTML = `<tr><td colspan="2" class="p-8 text-center text-red-600 font-bold">Failed to load: ${err.message}</td></tr>`;
          })
          .refreshData(['hrData']);
      }"""
    content = content.replace(target, replacement)

    with open('Index.html', 'w') as f:
        f.write(content)

if __name__ == "__main__":
    main()
