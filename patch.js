const fs = require('fs');
let html = fs.readFileSync('Index.html', 'utf8');

const injectionPoint = `function loadQuickCover() {`;

const newCode = `function getNext5SchoolDays() {
        const days = [];
        let current = new Date();
        current.setHours(12, 0, 0, 0); // Noon to avoid timezone issues

        while (days.length < 5) {
            const dayOfWeek = current.getDay();
            if (dayOfWeek !== 0 && dayOfWeek !== 6) { // Skip Sunday (0) and Saturday (6)
                const yyyy = current.getFullYear();
                const mm = String(current.getMonth() + 1).padStart(2, '0');
                const dd = String(current.getDate()).padStart(2, '0');
                days.push({
                    dateStr: \`\${yyyy}-\${mm}-\${dd}\`,
                    dateObj: new Date(current.getTime())
                });
            }
            current.setDate(current.getDate() + 1);
        }
        return days;
      }

      function renderSubAvailabilityDashboard() {
        const container = document.getElementById('subAvailabilityContainer');
        if (!staffList || staffList.length === 0) {
            container.innerHTML = '<div class="text-center text-gray-400 italic">No substitute data found.</div>';
            return;
        }

        const next5Days = getNext5SchoolDays();

        let html = '<div class="grid grid-cols-1 md:grid-cols-5 gap-4">';

        next5Days.forEach(dayInfo => {
            const dateStr = dayInfo.dateStr;
            const dateObj = dayInfo.dateObj;

            // Format date for display (e.g., "Mon, Oct 16")
            const displayDate = dateObj.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' });

            // Get day color
            const color = globalDateColors[dateStr] || "Unknown";
            let colorPill = '';
            if (color.toLowerCase() === 'green') {
                colorPill = '<span class="ml-2 inline-block bg-green-100 text-green-800 text-[10px] font-bold px-2 py-0.5 rounded border border-green-200 uppercase tracking-wider">Green</span>';
            } else if (color.toLowerCase() === 'blue') {
                colorPill = '<span class="ml-2 inline-block bg-blue-100 text-blue-800 text-[10px] font-bold px-2 py-0.5 rounded border border-blue-200 uppercase tracking-wider">Blue</span>';
            } else if (color.toLowerCase() === 'gold') {
                colorPill = '<span class="ml-2 inline-block bg-yellow-100 text-yellow-800 text-[10px] font-bold px-2 py-0.5 rounded border border-yellow-200 uppercase tracking-wider">Gold</span>';
            } else {
                colorPill = \`<span class="ml-2 inline-block bg-gray-100 text-gray-800 text-[10px] font-bold px-2 py-0.5 rounded border border-gray-200 uppercase tracking-wider">\${escapeHtml(color)}</span>\`;
            }

            // Find subs available on this day
            const availableAllDay = [];
            const availableAM = [];
            const availablePM = [];

            staffList.forEach(staff => {
                if (staff.role && staff.role.toLowerCase().includes('substitute')) {
                    const status = staff.availability && staff.availability[dateStr];
                    if (status === 'Available') {
                        availableAllDay.push(staff.name);
                    } else if (status === 'AM Only') {
                        availableAM.push(staff.name);
                    } else if (status === 'PM Only') {
                        availablePM.push(staff.name);
                    }
                }
            });

            // Build card for the day
            html += \`
                <div class="bg-gray-50 border border-gray-200 rounded-lg p-3 flex flex-col h-full">
                    <div class="flex items-center justify-between border-b border-gray-200 pb-2 mb-2">
                        <h4 class="font-bold text-[#002147] text-sm">\${escapeHtml(displayDate)}</h4>
                        \${colorPill}
                    </div>

                    <div class="flex-grow space-y-3 overflow-y-auto max-h-[300px]">
                        <div>
                            <h5 class="text-xs font-bold text-gray-500 uppercase tracking-wider mb-1">All Day (\${availableAllDay.length})</h5>
                            \${availableAllDay.length > 0 ?
                                '<ul class="text-sm text-green-700 font-medium list-disc list-inside">' + availableAllDay.map(name => \`<li>\${escapeHtml(name)}\</li>\`).join('') + '</ul>' :
                                '<div class="text-xs text-gray-400 italic">None</div>'}
                        </div>
                        <div>
                            <h5 class="text-xs font-bold text-gray-500 uppercase tracking-wider mb-1">AM Only (\${availableAM.length})</h5>
                            \${availableAM.length > 0 ?
                                '<ul class="text-sm text-blue-700 font-medium list-disc list-inside">' + availableAM.map(name => \`<li>\${escapeHtml(name)}\</li>\`).join('') + '</ul>' :
                                '<div class="text-xs text-gray-400 italic">None</div>'}
                        </div>
                        <div>
                            <h5 class="text-xs font-bold text-gray-500 uppercase tracking-wider mb-1">PM Only (\${availablePM.length})</h5>
                            \${availablePM.length > 0 ?
                                '<ul class="text-sm text-orange-700 font-medium list-disc list-inside">' + availablePM.map(name => \`<li>\${escapeHtml(name)}\</li>\`).join('') + '</ul>' :
                                '<div class="text-xs text-gray-400 italic">None</div>'}
                        </div>
                    </div>
                </div>
            \`;
        });

        html += '</div>';
        container.innerHTML = html;
        container.classList.remove('p-6');
        container.classList.add('p-4');
      }

      `;

if (html.includes(injectionPoint)) {
    html = html.replace(injectionPoint, newCode + injectionPoint);
    fs.writeFileSync('Index.html', html);
    console.log("Patched renderSubAvailabilityDashboard successfully");
} else {
    console.log("Could not find injection point");
}
