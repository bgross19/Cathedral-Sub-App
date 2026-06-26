import re

def main():
    with open('Index.html', 'r') as f:
        content = f.read()

    pubsub_code = """
      // --- PUB-SUB EVENT EMITTER ---
      const EventBus = {
        listeners: {},
        subscribe: function(event, callback) {
          if (!this.listeners[event]) {
            this.listeners[event] = [];
          }
          this.listeners[event].push(callback);
        },
        publish: function(event, data) {
          if (this.listeners[event]) {
            this.listeners[event].forEach(callback => callback(data));
          }
        }
      };

      // Subscribe UI components to data refresh events
      EventBus.subscribe('refresh_myAbsences', () => loadMyAbsences());
      EventBus.subscribe('refresh_mySubDuties', () => loadMySubDuties());
      EventBus.subscribe('refresh_todaysOpenJobs', () => loadTodaysOpenJobs());
      EventBus.subscribe('refresh_quickCover', () => loadQuickCover());
      EventBus.subscribe('refresh_adminData', () => { loadAdminDataRawSilent(); loadAdminDashboardData(); });
      // -------------------------------
"""

    # Inject PubSub at the beginning of the first <script> block
    # It's at `<script>\n      let currentAbsenceReasons` but wait, there are multiple <script> tags.
    # The main logic script is the second one starting with `function toggleSection`

    target_script = "<script>\n      function toggleSection"
    new_script = "<script>\n" + pubsub_code + "\n      function toggleSection"

    content = content.replace(target_script, new_script)

    # Now we need to update the data mutation functions to publish events instead of manually calling loadX().

    # 1. handleSignMeUp
    #  Original:
    #               loadTodaysOpenJobs();
    #               loadMySubDuties();
    #
    #               if (!document.getElementById('adminSection').classList.contains('hidden')) {
    #                  loadQuickCover();
    #                  loadAdminDataRawSilent();
    #               }
    #  We can replace this block with:
    #               EventBus.publish('refresh_todaysOpenJobs');
    #               EventBus.publish('refresh_mySubDuties');
    #               if (!document.getElementById('adminSection').classList.contains('hidden')) {
    #                  EventBus.publish('refresh_quickCover');
    #                  EventBus.publish('refresh_adminData');
    #               }

    handleSignMeUp_orig = """              loadTodaysOpenJobs();
              loadMySubDuties();

              if (!document.getElementById('adminSection').classList.contains('hidden')) {
                 loadQuickCover();
                 loadAdminDataRawSilent();
              }"""

    handleSignMeUp_new = """              EventBus.publish('refresh_todaysOpenJobs');
              EventBus.publish('refresh_mySubDuties');

              if (!document.getElementById('adminSection').classList.contains('hidden')) {
                 EventBus.publish('refresh_quickCover');
                 EventBus.publish('refresh_adminData');
              }"""

    content = content.replace(handleSignMeUp_orig, handleSignMeUp_new)

    # 2. handleAdminAssign
    #             // Check if At A Glance dashboard is currently visible
    #             const atAGlanceView = document.getElementById('atAGlanceDashboardView');
    #             const isAtAGlanceVisible = atAGlanceView && !atAGlanceView.classList.contains('hidden');
    #
    #             // Refresh admin data behind the scenes
    #             loadAdminDashboardData(isAtAGlanceVisible);
    #
    #             // Also refresh quick cover if it's the same 2-day window
    #             loadQuickCover();
    #             loadAdminDataRawSilent();
    #
    # Wait, EventBus.publish('refresh_adminData') calls both raw silent and dashboard.
    # But admin dashboard needs the `isAtAGlanceVisible` flag.
    # Let's pass data to the event!
    # EventBus.publish('refresh_adminData', { renderAtAGlance: isAtAGlanceVisible });
    # And modify subscribe: EventBus.subscribe('refresh_adminData', (data) => { loadAdminDataRawSilent(); loadAdminDashboardData(data && data.renderAtAGlance); });

    pubsub_code_new = pubsub_code.replace("EventBus.subscribe('refresh_adminData', () => { loadAdminDataRawSilent(); loadAdminDashboardData(); });",
                                          "EventBus.subscribe('refresh_adminData', (data) => { loadAdminDataRawSilent(); loadAdminDashboardData(data && data.renderAtAGlance); });")
    content = content.replace(pubsub_code, pubsub_code_new)

    handleAdminAssign_orig = """              // Refresh admin data behind the scenes
              loadAdminDashboardData(isAtAGlanceVisible);

              // Also refresh quick cover if it's the same 2-day window
              loadQuickCover();
              loadAdminDataRawSilent();"""

    handleAdminAssign_new = """              // Refresh admin data behind the scenes
              EventBus.publish('refresh_adminData', { renderAtAGlance: isAtAGlanceVisible });

              // Also refresh quick cover if it's the same 2-day window
              EventBus.publish('refresh_quickCover');"""

    content = content.replace(handleAdminAssign_orig, handleAdminAssign_new)

    # 3. handleAssign (from quick cover)
    #               loadQuickCover(); // Refresh the list
    #               loadAdminDataRawSilent();
    handleAssign_orig = """              loadQuickCover(); // Refresh the list
              loadAdminDataRawSilent();"""
    handleAssign_new = """              EventBus.publish('refresh_quickCover'); // Refresh the list
              EventBus.publish('refresh_adminData');"""
    content = content.replace(handleAssign_orig, handleAssign_new)

    # 4. cancelMySubDutyClient
    #                loadMySubDuties();
    #                if (!document.getElementById('adminSection').classList.contains('hidden')) {
    #                  loadQuickCover();
    #                  loadAdminDataRawSilent();
    #                }
    cancelDuty_orig = """               loadMySubDuties();
               if (!document.getElementById('adminSection').classList.contains('hidden')) {
                 loadQuickCover();
                 loadAdminDataRawSilent();
               }"""
    cancelDuty_new = """               EventBus.publish('refresh_mySubDuties');
               if (!document.getElementById('adminSection').classList.contains('hidden')) {
                 EventBus.publish('refresh_quickCover');
                 EventBus.publish('refresh_adminData');
               }"""
    content = content.replace(cancelDuty_orig, cancelDuty_new)

    # 5. deleteAbsenceFromModal
    #                 if (!document.getElementById('adminSection').classList.contains('hidden')) {
    #                   loadQuickCover();
    #                   loadAdminDataRawSilent();
    #                   // Admin dashboard view handles its own UI refresh when raw data is reloaded, but let's force a refresh
    #                   loadAdminDashboardData();
    #                 }
    deleteAbs_orig = """                if (!document.getElementById('adminSection').classList.contains('hidden')) {
                  loadQuickCover();
                  loadAdminDataRawSilent();
                  // Admin dashboard view handles its own UI refresh when raw data is reloaded, but let's force a refresh
                  loadAdminDashboardData();
                }"""
    deleteAbs_new = """                if (!document.getElementById('adminSection').classList.contains('hidden')) {
                  EventBus.publish('refresh_quickCover');
                  EventBus.publish('refresh_adminData');
                }"""
    content = content.replace(deleteAbs_orig, deleteAbs_new)

    # 6. cancelAbsence
    #                loadMyAbsences();
    #                if (!document.getElementById('adminSection').classList.contains('hidden')) {
    #                  loadQuickCover();
    #                  loadAdminDataRawSilent();
    #                }
    cancelAbs_orig = """               loadMyAbsences();
               if (!document.getElementById('adminSection').classList.contains('hidden')) {
                 loadQuickCover();
                 loadAdminDataRawSilent();
               }"""
    cancelAbs_new = """               EventBus.publish('refresh_myAbsences');
               if (!document.getElementById('adminSection').classList.contains('hidden')) {
                 EventBus.publish('refresh_quickCover');
                 EventBus.publish('refresh_adminData');
               }"""
    content = content.replace(cancelAbs_orig, cancelAbs_new)

    # 7. Edit/Submit inside absenceForm.onsubmit
    #                 loadMyAbsences();
    #                 if (!document.getElementById('adminSection').classList.contains('hidden')) {
    #                   loadQuickCover();
    #                   loadAdminDataRawSilent();
    #                   // Admin dashboard view handles its own UI refresh when raw data is reloaded, but let's force a refresh
    #                   loadAdminDashboardData();
    #                 }
    onsubmit_edit_orig = """                loadMyAbsences();
                if (!document.getElementById('adminSection').classList.contains('hidden')) {
                  loadQuickCover();
                  loadAdminDataRawSilent();
                  // Admin dashboard view handles its own UI refresh when raw data is reloaded, but let's force a refresh
                  loadAdminDashboardData();
                }"""
    onsubmit_edit_new = """                EventBus.publish('refresh_myAbsences');
                if (!document.getElementById('adminSection').classList.contains('hidden')) {
                  EventBus.publish('refresh_quickCover');
                  EventBus.publish('refresh_adminData');
                }"""
    content = content.replace(onsubmit_edit_orig, onsubmit_edit_new)

    # Submit new
    #                 loadMyAbsences();
    #
    #                 if (!document.getElementById('adminSection').classList.contains('hidden')) {
    #                   loadQuickCover();
    #                   loadAdminDataRawSilent();
    #                 }
    onsubmit_new_orig = """                loadMyAbsences();

                if (!document.getElementById('adminSection').classList.contains('hidden')) {
                  loadQuickCover();
                  loadAdminDataRawSilent();
                }"""
    onsubmit_new_new = """                EventBus.publish('refresh_myAbsences');

                if (!document.getElementById('adminSection').classList.contains('hidden')) {
                  EventBus.publish('refresh_quickCover');
                  EventBus.publish('refresh_adminData');
                }"""
    content = content.replace(onsubmit_new_orig, onsubmit_new_new)

    with open('Index.html', 'w') as f:
        f.write(content)

if __name__ == "__main__":
    main()
