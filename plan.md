1. **Update `Index.html` to add the "Add Availability for Sub" button and modal.**
    *   Add a new button next to "Add Request for Teacher" that is visible only to users with the `Add Request on Behalf` permission.
    *   Add a new modal (or a new section in an existing modal) for selecting a substitute and editing their availability.
    *   Add a `datalist` of substitutes for the modal.
    *   Add a calendar view in the modal, similar to the "My Availability" calendar.

2. **Add JavaScript functions to handle the new modal.**
    *   `toggleSubAvailabilityModal()` to open/close the modal.
    *   `handleOnBehalfSubChange()` to fetch and display the selected substitute's availability.
    *   `toggleAvailabilityForSub(dateStr)` to toggle availability for the selected substitute.
    *   `renderAvailabilityCalendarForSub()` to render the calendar in the modal.
    *   `initAvailabilityCalendarForSub(initialData)` to initialize the calendar data.
    *   `changeAvailabilityMonthForSub(delta)` to change the month being viewed.

3. **Update backend functions in `code.gs`.**
    *   Create a new backend function `getSubstituteAvailabilityForAdmin(targetEmail, clientEmail)` that takes the target substitute's email and returns their availability (only if the `clientEmail` has the correct permission). Or simply reuse `getSubstituteAvailability` if we just need to pass the target email.
    *   Update `saveSubstituteAvailability` to optionally take a `targetEmail` parameter for admins editing on behalf, or create a new function `saveSubstituteAvailabilityAdmin(targetEmail, dateStr, status, clientEmail)`.

4. **Verify permissions.**
    *   Make sure only users with `Add Request on Behalf` permission can see the button and use the backend functions for other users.

5. **Pre-commit checks.**
    *   Run `pre_commit_instructions` to ensure all necessary steps are completed.
