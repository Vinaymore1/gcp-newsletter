// Toggle expandable sections
    function toggleDetails(button) {
        const card = button.closest('.expandable');
        const wasExpanded = card.classList.contains('expanded');
        
        // Close all other expanded sections first
        document.querySelectorAll('.expandable.expanded').forEach(el => {
            if (el !== card) {
                el.classList.remove('expanded');
            }
        });
        
        // Toggle the clicked section
        card.classList.toggle('expanded', !wasExpanded);
    }

    // Close expanded sections
    function closeDetails() {
        document.querySelectorAll('.expandable').forEach(el => {
            el.classList.remove('expanded');
        });
    }

    // Toggle security alert details
    function toggleSecurityDetails(button) {
        const alert = button.closest('.security-alert');
        const viewBtn = alert.querySelector('#view');
        const closeBtn = alert.querySelector('#close');
        const wasExpanded = alert.classList.contains('expanded');
        
        // Toggle button visibility
        viewBtn.style.display = wasExpanded ? 'block' : 'none';
        closeBtn.style.display = wasExpanded ? 'none' : 'block';
        
        // Toggle expanded state
        alert.classList.toggle('expanded', !wasExpanded);
    }

    // Close security alert details
    function closeSecurityDetails(button) {
        const alert = button.closest('.security-alert');
        const viewBtn = alert.querySelector('#view');
        const closeBtn = alert.querySelector('#close');
        
        // Toggle button visibility
        viewBtn.style.display = 'block';
        closeBtn.style.display = 'none';
        
        // Remove expanded state
        alert.classList.remove('expanded');
    }