// ==UserScript==
// @name         VoiceWSS Contact Autofill
// @namespace    http://tampermonkey.net/
// @version      3.3
// @description  Extract names and phone numbers from WebSelfStorage reports and autofill on Google Voice
// @author       You
// @match        https://voice.google.com/*
// @match        https://messages.google.com/web/*
// @match        https://webselfstorage.com/Affiliate/*/Reports/ViewReport?reportID=*
// @grant        GM_setValue
// @grant        GM_getValue
// @updateURL    https://raw.githubusercontent.com/33OPX/VoiceWSS-Contact-Autofill/main/Google%20Voice%20WSS.js
// @downloadURL  https://raw.githubusercontent.com/33OPX/VoiceWSS-Contact-Autofill/main/Google%20Voice%20WSS.js
// ==/UserScript==


(function() {
    // Preload default message templates on script load if missing
    (function ensureDefaultTemplates() {
        let stored = {};
        try { stored = JSON.parse(GM_getValue('wss_msg_templates', '{}')); } catch (e) { stored = {}; }
        if (!stored || Object.keys(stored).length === 0) {
            const defaultTemplates = {
                stdP1: 'This is <employeename> from the U-Haul. I’m reaching out because I really want to help you reclaim your belongings, and you have options. The next step is your unit being listed for auction, and I REALLY don’t want that to happen. What can I do to help? Are you interested in your belongings?',
                stdP2: 'Hi <firstname>, this is <employeename> from U-Haul. Just checking in to see if you need anything or have questions about your storage unit.',
                stdA1: 'Hey <altfirstname>, my name is <employeename> and I’m the Storage Manager for U-Haul in this area. I’m reaching out because <firstname> has listed you as their emergency contact for their storage unit. We really need to get in contact with them regarding their belongings. Do you have a good number for us to get ahold of them at? Thank you so much!',
                stdA2: 'Hi <altfirstname>, this is <employeename> from U-Haul. Just checking in to see if you need anything or have questions about <firstname>\'s storage unit.',
                aucP1: 'Hey <firstname>, my name is <employeename> and I’m the Storage Manager for U-Haul in this area. We are desperately trying to get in contact with our customer regarding their belongings. Is this by chance your storage unit? Please let me know either way and thank you so much for your help!',
                aucP2: 'Hi <firstname>, this is <employeename> from U-Haul. Just checking in to see if you need anything or have questions about your storage unit.',
                aucA1: 'Hey <altfirstname>, my name is <employeename> and I’m the Storage Manager for U-Haul in this area. I’m reaching out because <firstname> has listed you as their emergency contact for their storage unit. We really need to get in contact with them regarding their belongings. Do you have a good number for us to get ahold of them at? Thank you so much!',
                aucA2: 'Hi <altfirstname>, this is <employeename> from U-Haul. Just checking in to see if you need anything or have questions about <firstname>\'s storage unit.',
                standard_daysMin: 29,
                standard_daysMax: 59,
                auction_daysMin: 60,
                auction_daysMax: 999
            };
            GM_setValue('wss_msg_templates', JSON.stringify(defaultTemplates));
        }
    })();
    'use strict';

    // Utility: Extract contacts from report page
    function extractContacts() {
        let contacts = [];
        // Find all rows in the report table
        const tables = document.querySelectorAll('table.report');
        tables.forEach(table => {
            const rows = Array.from(table.querySelectorAll('tr'));
            let daysLate = null;
            let balanceDue = null;
            let unitNumber = null;
            let lastMainContact = null;
            for (let i = 0; i < rows.length; i++) {
                const rowText = rows[i].innerText.trim();
                // Find days late, balance due, and unit number from the main data row
                if (/^\d{6,}-\d{6,}/.test(rowText)) {
                    let cols = rowText.split('\t');
                    // Example columns: [unit, ?, ?, ?, daysLate, balanceDue, ...]
                    if (cols.length >= 6) {
                        // Try to get unit number (first column)
                        unitNumber = cols[0].trim();
                        // Try to get days late (5th column)
                        if (/^\d+$/.test(cols[4].trim())) {
                            daysLate = parseInt(cols[4].trim(), 10);
                        } else {
                            daysLate = null;
                        }
                        // Try to get balance due (6th column)
                        let bal = cols[5].replace(/[^\d\.\-]/g, '').trim();
                        balanceDue = bal.length > 0 ? bal : null;
                    } else {
                        daysLate = null;
                        balanceDue = null;
                        unitNumber = null;
                    }
                }
                // Main customer
                if (/^Customer:/.test(rowText)) {
                    let nameMatch = rowText.match(/Customer:\s*([A-Z\s]+,[A-Z\s]+)/);
                    if (nameMatch) {
                        let nameText = nameMatch[1].trim();
                        // Remove all AUTOPAY and N, and clean up tabs/whitespace
                        nameText = nameText.replace(/\t+/g, ' ')
                                         .replace(/\bAUTOPAY\b/gi, '')
                                         .replace(/\bN\b/gi, '')
                                         .replace(/\s+/g, ' ')
                                         .trim();
                        // Look ahead for Home phone
                        let phone = null;
                        for (let j = 1; j <= 3 && i + j < rows.length; j++) {
                            let phoneRow = rows[i + j].innerText.trim();
                            let phoneMatch = phoneRow.match(/Home\s*:\s*(\(?\d{3}\)?[\-\s]?\d{3}[\-\s]?\d{4})/);
                            if (phoneMatch) {
                                phone = phoneMatch[1];
                                break;
                            }
                        }
                        let contact = {
                            name: nameText,
                            phone,
                            daysLate,
                            alternates: [],
                            balanceDue,
                            unit: unitNumber
                        };
                        contacts.push(contact);
                        lastMainContact = contact;
                    }
                }
                // Alternate contact
                if (/^Alternate Contact/.test(rowText)) {
                    // Look ahead for name and phone
                    let altName = null, altPhone = null;
                    for (let j = 1; j <= 3 && i + j < rows.length; j++) {
                        let altRow = rows[i + j].innerText.trim();
                        let nameMatch = altRow.match(/^([A-Z\s]+,[A-Z\s]+)/);
                        if (nameMatch) {
                            altName = nameMatch[1].trim();
                        }
                        let phoneMatch = altRow.match(/(\(?\d{3}\)?[\-\s]?\d{3}[\-\s]?\d{4})/);
                        if (phoneMatch) {
                            altPhone = phoneMatch[1];
                        }
                        if (altName && altPhone) break;
                    }
                    if (altName && altPhone && lastMainContact) {
                        // Pass balanceDue and unit to alternates as well
                        lastMainContact.alternates.push({
                            name: altName,
                            phone: altPhone,
                            daysLate,
                            balanceDue: lastMainContact.balanceDue,
                            unit: lastMainContact.unit
                        });
                    }
                }
            }
        });
        return contacts;
    }

    // Save contacts to Tampermonkey cross-domain storage
    function saveContacts(contacts) {
        GM_setValue('wss_contacts', JSON.stringify(contacts));
    }

    // Load contacts from Tampermonkey cross-domain storage
    function loadContacts() {
        let data = GM_getValue('wss_contacts', '[]');
        console.log('VoiceWSS: Raw GM wss_contacts:', data);
        let parsed = [];
        try {
            parsed = data ? JSON.parse(data) : [];
        } catch (e) {
            console.error('VoiceWSS: Error parsing contacts from GM storage', e);
        }
        console.log('VoiceWSS: Parsed contacts:', parsed);
        return parsed;
    }

    // On WebSelfStorage report page: extract and save contacts
    if (window.location.hostname.includes('webselfstorage.com')) {
        let contacts = extractContacts();
        // Always clear and replace previous list
        saveContacts([]); // Clear previous list
        if (contacts.length > 0) {
            saveContacts(contacts);
            alert(`Saved ${contacts.length} contacts for Google Voice.`);
        } else {
            alert('No contacts found to save.');
        }
    }

    // On Google Voice or Google Messages Web: inject dropdown at the top of the page
    function injectDropdown() {
        if (document.querySelector('.wss-contact-dropdown')) return; // Already injected
        console.log('VoiceWSS: Injecting contacts dropdown');
        let contacts = loadContacts();
        // Sort contacts by daysLate descending (highest to lowest)
        contacts.sort((a, b) => {
            // If daysLate is missing, treat as 0
            let aLate = (typeof a.daysLate === 'number' && !isNaN(a.daysLate)) ? a.daysLate : 0;
            let bLate = (typeof b.daysLate === 'number' && !isNaN(b.daysLate)) ? b.daysLate : 0;
            return bLate - aLate;
        });
        let container = document.createElement('div');
    container.className = 'wss-contact-dropdown';
    container.style.background = 'linear-gradient(135deg, #e6f2ff 0%, #f8faff 100%)';
    container.style.padding = '5px 18px';
    container.style.borderRadius = '12px';
    container.style.border = '1px solid #0077cc';
    container.style.zIndex = '9999';
    container.style.position = 'absolute';
    container.style.left = '32px';
    container.style.top = '32px';
    container.style.maxWidth = '295px';
    container.style.width = '295px';
    container.style.boxSizing = 'border-box';
    container.style.textAlign = 'center';
    container.style.display = 'block';
    container.style.boxShadow = '0 2px 12px rgba(0,64,128,0.10)';

    // Load saved position from Tampermonkey storage
    let savedPos = {};
    try { savedPos = JSON.parse(GM_getValue('wss_dropdown_pos', '{}')); } catch (e) { savedPos = {}; }
    if (typeof savedPos.left === 'number' && typeof savedPos.top === 'number') {
        container.style.left = savedPos.left + 'px';
        container.style.top = savedPos.top + 'px';
    }

    // Add draggable handle (4-way arrow)
    let dragHandle = document.createElement('div');
    dragHandle.textContent = '✥'; // Four-pointed star symbol
    dragHandle.title = 'Drag to move';
    dragHandle.style.position = 'absolute';
    dragHandle.style.left = '6px';
    dragHandle.style.top = '6px';
    dragHandle.style.cursor = 'grab';
    dragHandle.style.fontSize = '22px';
    dragHandle.style.userSelect = 'none';
    dragHandle.style.zIndex = '10001';
    dragHandle.style.background = 'transparent';
    dragHandle.style.color = '#0077cc';
    dragHandle.style.padding = '0 2px';
    dragHandle.style.borderRadius = '6px';
    dragHandle.onmousedown = function(e) {
        e.preventDefault();
        let startX = e.clientX;
        let startY = e.clientY;
        let origLeft = parseInt(container.style.left, 10);
        let origTop = parseInt(container.style.top, 10);
        dragHandle.style.cursor = 'grabbing';
        function onMouseMove(ev) {
            let dx = ev.clientX - startX;
            let dy = ev.clientY - startY;
            let newLeft = origLeft + dx;
            let newTop = origTop + dy;
            container.style.left = newLeft + 'px';
            container.style.top = newTop + 'px';
            // Save position live as you drag
            GM_setValue('wss_dropdown_pos', JSON.stringify({ left: newLeft, top: newTop }));
        }
        function onMouseUp() {
            dragHandle.style.cursor = 'grab';
            document.removeEventListener('mousemove', onMouseMove);
            document.removeEventListener('mouseup', onMouseUp);
        }
        document.addEventListener('mousemove', onMouseMove);
        document.addEventListener('mouseup', onMouseUp);
    };
    container.appendChild(dragHandle);
    // Position left for Google Voice, centered for Google Messages
    if (window.location.hostname.includes('voice.google.com')) {
        container.style.margin = '16px 0 0 16px';
        container.style.float = 'left';
    } else {
        container.style.margin = '0 auto';
        container.style.float = '';
    }

        // If on Google Messages, wrap in a flexbox container to prevent stretching
        let wrapper = null;
        // --- Refresh Button ---
        let refreshBtn = document.createElement('button');
    refreshBtn.textContent = '🔄';
        refreshBtn.style.marginLeft = '8px';
        refreshBtn.style.marginRight = '0px';
        refreshBtn.style.padding = '2px 8px';
        refreshBtn.style.borderRadius = '5px';
    refreshBtn.style.border = 'none';
        refreshBtn.style.background = 'linear-gradient(90deg, #e6f2ff 0%, #f8faff 100%)';
        refreshBtn.style.cursor = 'pointer';
        refreshBtn.style.fontWeight = '500';
    refreshBtn.style.fontSize = '20px';
        refreshBtn.style.boxShadow = '0 1px 4px rgba(0,0,0,0.07)';
        refreshBtn.style.whiteSpace = 'nowrap';
        refreshBtn.style.lineHeight = '1.2';
        refreshBtn.style.height = 'auto';
        refreshBtn.onmouseover = function() { refreshBtn.style.background = '#d0e7ff'; };
        refreshBtn.onmouseout = function() { refreshBtn.style.background = 'linear-gradient(90deg, #e6f2ff 0%, #f8faff 100%)'; };
        refreshBtn.onclick = function() {
            // Remove the dropdown container and re-inject everything
            let dd = document.querySelector('.wss-contact-dropdown');
            if (dd && dd.parentNode) dd.parentNode.removeChild(dd);
            injectDropdown();
        };
        // --- SMS Editor UI ---
        // Create SMS Editor button
        let smsEditorBtn = document.createElement('button');
        smsEditorBtn.textContent = '✏️ SMS Editor';
        smsEditorBtn.style.marginLeft = '4px';
        smsEditorBtn.style.marginRight = '0px';
        smsEditorBtn.style.padding = '2px 6px';
        smsEditorBtn.style.borderRadius = '5px';
        smsEditorBtn.style.border = '1px solid #0077cc';
        smsEditorBtn.style.background = 'linear-gradient(90deg, #e6f2ff 0%, #f8faff 100%)';
        smsEditorBtn.style.cursor = 'pointer';
        smsEditorBtn.style.fontWeight = '500';
        smsEditorBtn.style.fontSize = '13px';
        smsEditorBtn.style.boxShadow = '0 1px 4px rgba(0,0,0,0.07)';
        smsEditorBtn.style.whiteSpace = 'nowrap';
        smsEditorBtn.style.lineHeight = '1.2';
        smsEditorBtn.style.height = 'auto';
        smsEditorBtn.onmouseover = function() { smsEditorBtn.style.background = '#d0e7ff'; };
        smsEditorBtn.onmouseout = function() { smsEditorBtn.style.background = 'linear-gradient(90deg, #e6f2ff 0%, #f8faff 100%)'; };
        // Editor panel (hidden by default)
        let smsEditorPanel = document.createElement('div');
    smsEditorPanel.style.position = 'fixed';
    smsEditorPanel.style.top = '60px';
    smsEditorPanel.style.left = '50%';
    smsEditorPanel.style.transform = 'translateX(-50%)';
    smsEditorPanel.style.zIndex = '100000';
    smsEditorPanel.style.background = 'linear-gradient(135deg, #e6f2ff 0%, #f8faff 100%)';
    smsEditorPanel.style.border = '1px solid #0077cc';
    smsEditorPanel.style.borderRadius = '12px';
    smsEditorPanel.style.boxShadow = '0 2px 12px rgba(0,64,128,0.10)';
    smsEditorPanel.style.padding = '12px 18px';
    smsEditorPanel.style.minWidth = '340px';
    smsEditorPanel.style.display = 'none';
    smsEditorPanel.style.maxWidth = '95vw';
    smsEditorPanel.style.fontFamily = 'inherit';
    smsEditorPanel.style.fontSize = '14px';

        // Message types and config
        const defaultTemplates = {
            main: {
                label: 'Standard (29-59 days late)',
                key: 'main',
                daysMin: 29,
                daysMax: 59,
                text: 'This is <employeename> from the U-Haul. I’m reaching out because I really want to help you reclaim your belongings, and you have options. The next step is your unit being listed for auction, and I REALLY don’t want that to happen. What can I do to help? Are you interested in your belongings?'
            },
            auction: {
                label: 'Auction (60+ days late)',
                key: 'auction',
                daysMin: 60,
                daysMax: 999,
                text: 'Hey <firstname>, my name is <employeename> and I’m the Storage Manager for U-Haul in this area. We are desperately trying to get in contact with our customer regarding their belongings. Is this by chance your storage unit? Please let me know either way and thank you so much for your help!'
            },
            alt: {
                label: 'Alternate Contact',
                key: 'alt',
                daysMin: 0,
                daysMax: 999,
                text: 'Hey <altfirstname>, my name is <employeename> and I’m the Storage Manager for U-Haul in this area. I’m reaching out because <firstname> has listed you as their emergency contact for their storage unit. We really need to get in contact with them regarding their belongings. Do you have a good number for us to get ahold of them at? Thank you so much!'
            },
            secondary: {
                label: 'Secondary (Alternate Type)',
                key: 'secondary',
                daysMin: 0,
                daysMax: 999,
                text: 'Hi <firstname>, this is <employeename> from U-Haul. Just checking in to see if you need anything or have questions about your storage unit.'
            }
        };

        // Load templates from storage or use defaults
        function loadTemplates() {
            let stored = {};
            try { stored = JSON.parse(GM_getValue('wss_msg_templates', '{}')); } catch (e) { stored = {}; }
            // If no templates exist, preload defaults
            if (!stored || Object.keys(stored).length === 0) {
                let preload = {
                    stdP1: defaultTemplates.main.text,
                    stdP2: defaultTemplates.secondary.text,
                    stdA1: defaultTemplates.alt.text,
                    stdA2: defaultTemplates.secondary.text,
                    aucP1: defaultTemplates.auction.text,
                    aucP2: defaultTemplates.secondary.text,
                    aucA1: defaultTemplates.alt.text,
                    aucA2: defaultTemplates.secondary.text,
                    standard_daysMin: defaultTemplates.main.daysMin,
                    standard_daysMax: defaultTemplates.main.daysMax,
                    auction_daysMin: defaultTemplates.auction.daysMin,
                    auction_daysMax: defaultTemplates.auction.daysMax
                };
                GM_setValue('wss_msg_templates', JSON.stringify(preload));
                stored = preload;
            }
            let templates = {};
            for (let k in defaultTemplates) {
                // Use correct mapping for UI
                let key = '';
                if (k === 'main') key = 'stdP1';
                else if (k === 'secondary') key = 'stdP2';
                else if (k === 'alt') key = 'stdA1';
                else if (k === 'auction') key = 'aucP1';
                templates[k] = Object.assign({}, defaultTemplates[k], stored[key] ? { text: stored[key] } : {});
            }
            return templates;
        }
        function saveTemplates(templates) {
            GM_setValue('wss_msg_templates', JSON.stringify(templates));
        }

        // Placeholders
        const placeholders = [
            { label: '<employeename>', desc: 'Employee Name' },
            { label: '<firstname>', desc: 'Customer First Name' },
            { label: '<altfirstname>', desc: 'Alternate Contact First Name' }
        ];

        // Build editor UI
        function showEditor() {
            // --- Custom Ranges Management ---
            // Load custom ranges from storage
            function loadCustomRanges() {
                let ranges = [];
                try { ranges = JSON.parse(GM_getValue('wss_custom_ranges', '[]')); } catch (e) { ranges = []; }
                return Array.isArray(ranges) ? ranges : [];
            }
            function saveCustomRanges(ranges) {
                GM_setValue('wss_custom_ranges', JSON.stringify(ranges));
            }
            let customRanges = loadCustomRanges();

            // Remove all children safely to avoid TrustedHTML CSP error
            while (smsEditorPanel.firstChild) {
                smsEditorPanel.removeChild(smsEditorPanel.firstChild);
            }
            // Add button to show/hide custom ranges UI
            let customRangesBtn = document.createElement('button');
            customRangesBtn.textContent = 'Add/Edit Custom Ranges';
            customRangesBtn.style.marginBottom = '10px';
            customRangesBtn.style.padding = '4px 12px';
            customRangesBtn.style.borderRadius = '6px';
            customRangesBtn.style.border = '1px solid #0077cc';
            customRangesBtn.style.background = '#e6f2ff';
            customRangesBtn.style.cursor = 'pointer';
            customRangesBtn.style.fontWeight = 'bold';
            smsEditorPanel.appendChild(customRangesBtn);

            // Custom ranges section (hidden by default)
            let customRangeSection = document.createElement('div');
            customRangeSection.style.marginBottom = '12px';
            customRangeSection.style.background = '#f5faff';
            customRangeSection.style.padding = '8px';
            customRangeSection.style.borderRadius = '6px';
            customRangeSection.style.border = '1px solid #b3d1ff';
            customRangeSection.style.fontSize = '14px';
            customRangeSection.style.display = 'none';
            let customTitle = document.createElement('b');
            customTitle.textContent = 'Custom Ranges';
            customRangeSection.appendChild(customTitle);

            // List custom ranges
            let rangeList = document.createElement('ul');
            rangeList.style.listStyle = 'none';
            rangeList.style.padding = '0';
            customRanges.forEach((r, idx) => {
                let li = document.createElement('li');
                li.style.marginBottom = '4px';
                li.style.display = 'flex';
                li.style.alignItems = 'center';
                let label = document.createElement('span');
                label.textContent = `${r.label} (${r.daysMin}-${r.daysMax} days)`;
                label.style.flex = '1';
                li.appendChild(label);
                // Edit button
                let editBtn = document.createElement('button');
                editBtn.textContent = 'Edit';
                editBtn.style.marginRight = '6px';
                editBtn.style.fontSize = '12px';
                editBtn.onclick = function() {
                    rangeLabelInput.value = r.label;
                    rangeMinInput.value = r.daysMin;
                    rangeMaxInput.value = r.daysMax;
                    editingIdx = idx;
                    addBtn.textContent = 'Update';
                };
                li.appendChild(editBtn);
                // Delete button
                let delBtn = document.createElement('button');
                delBtn.textContent = 'Delete';
                delBtn.style.fontSize = '12px';
                delBtn.onclick = function() {
                    if (confirm('Delete this custom range?')) {
                        customRanges.splice(idx, 1);
                        saveCustomRanges(customRanges);
                        showEditor();
                    }
                };
                li.appendChild(delBtn);
                rangeList.appendChild(li);
            });
            customRangeSection.appendChild(rangeList);

            // Add/edit form
            let formRow = document.createElement('div');
            formRow.style.display = 'flex';
            formRow.style.gap = '6px';
            formRow.style.marginTop = '6px';
            let rangeLabelInput = document.createElement('input');
            rangeLabelInput.type = 'text';
            rangeLabelInput.placeholder = 'Label';
            rangeLabelInput.style.width = '90px';
            let rangeMinInput = document.createElement('input');
            rangeMinInput.type = 'number';
            rangeMinInput.placeholder = 'Min';
            rangeMinInput.style.width = '60px';
            let rangeMaxInput = document.createElement('input');
            rangeMaxInput.type = 'number';
            rangeMaxInput.placeholder = 'Max';
            rangeMaxInput.style.width = '60px';
            let addBtn = document.createElement('button');
            addBtn.textContent = 'Add';
            let editingIdx = null;
            addBtn.onclick = function() {
                let label = rangeLabelInput.value.trim();
                let min = parseInt(rangeMinInput.value, 10);
                let max = parseInt(rangeMaxInput.value, 10);
                if (!label || isNaN(min) || isNaN(max) || min > max) {
                    alert('Please enter valid label and min/max days.');
                    return;
                }
                if (editingIdx !== null) {
                    customRanges[editingIdx] = { label, daysMin: min, daysMax: max };
                    editingIdx = null;
                    addBtn.textContent = 'Add';
                } else {
                    customRanges.push({ label, daysMin: min, daysMax: max });
                }
                saveCustomRanges(customRanges);
                rangeLabelInput.value = '';
                rangeMinInput.value = '';
                rangeMaxInput.value = '';
                showEditor();
            };
            formRow.appendChild(rangeLabelInput);
            formRow.appendChild(rangeMinInput);
            formRow.appendChild(rangeMaxInput);
            formRow.appendChild(addBtn);
            customRangeSection.appendChild(formRow);

            // Show/hide logic
            customRangesBtn.onclick = function() {
                customRangeSection.style.display = (customRangeSection.style.display === 'none') ? 'block' : 'none';
            };
            smsEditorPanel.appendChild(customRangeSection);
            // Eight templates: stdP1, stdP2, stdA1, stdA2, aucP1, aucP2, aucA1, aucA2
            let templates = loadTemplates();
            let selectedDropdown = 'standard'; // 'standard' or 'auction'
            let selectedButton = 0; // 0: P1, 1: P2, 2: A1, 3: A2

            // Always load min/max from storage for dropdown labels
            let storedTemplates = {};
            try { storedTemplates = JSON.parse(GM_getValue('wss_msg_templates', '{}')); } catch (e) { storedTemplates = {}; }
            let stdMin = storedTemplates['standard_daysMin'] !== undefined ? storedTemplates['standard_daysMin'] : 29;
            let stdMax = storedTemplates['standard_daysMax'] !== undefined ? storedTemplates['standard_daysMax'] : 59;
            let aucMin = storedTemplates['auction_daysMin'] !== undefined ? storedTemplates['auction_daysMin'] : 60;
            let aucMax = storedTemplates['auction_daysMax'] !== undefined ? storedTemplates['auction_daysMax'] : 999;
            // Build all ranges into a sortable array
            let allRanges = [
                { value: 'standard', label: `Standard (days ${stdMin}-${stdMax})`, daysMin: stdMin, daysMax: stdMax, type: 'standard' },
                { value: 'auction', label: `Auction (${aucMin}+ Days)`, daysMin: aucMin, daysMax: aucMax, type: 'auction' }
            ];
            customRanges.forEach((r, idx) => {
                allRanges.push({ value: 'custom_' + idx, label: `${r.label} (${r.daysMin}-${r.daysMax} days)`, daysMin: r.daysMin, daysMax: r.daysMax, type: 'custom', idx });
            });
            // Sort by daysMin ascending
            allRanges.sort((a, b) => a.daysMin - b.daysMin);
            // Dropdown for days-late range
            let templateDropdown = document.createElement('select');
            templateDropdown.style.marginBottom = '12px';
            templateDropdown.style.width = '100%';
            let stdOpt = null, auctionOpt = null;
            allRanges.forEach(r => {
                let opt = document.createElement('option');
                opt.value = r.value;
                opt.textContent = r.label;
                templateDropdown.appendChild(opt);
                if (r.value === 'standard') stdOpt = opt;
                if (r.value === 'auction') auctionOpt = opt;
            });
            templateDropdown.value = selectedDropdown;
            smsEditorPanel.appendChild(templateDropdown);

            // Button row: SMS 1/2 Primary/Alternate
            let typeRow = document.createElement('div');
            typeRow.style.display = 'flex';
            typeRow.style.gap = '8px';
            typeRow.style.marginBottom = '12px';
            const buttonList = [
                { label: 'SMS 1 (Primary)', idx: 0 },
                { label: 'SMS 2 (Primary)', idx: 1 },
                { label: 'SMS 1 (Alternate)', idx: 2 },
                { label: 'SMS 2 (Alternate)', idx: 3 }
            ];
            buttonList.forEach((b, idx) => {
                let btn = document.createElement('button');
                btn.className = 'uh-copy-btn';
                btn.innerText = b.label;
                btn.style.fontWeight = '500';
                btn.style.fontSize = '13px';
                btn.style.padding = '2px 8px';
                btn.style.borderRadius = '5px';
                btn.style.border = '1px solid #0077cc';
                btn.style.background = (selectedButton === idx) ? 'linear-gradient(90deg, #e6f2ff 0%, #f8faff 100%)' : '#f8f8f8';
                btn.style.margin = '0 2px';
                btn.style.boxShadow = '0 1px 4px rgba(0,0,0,0.07)';
                btn.style.cursor = 'pointer';
                btn.onmouseover = function() { btn.style.background = '#d0e7ff'; };
                btn.onmouseout = function() { btn.style.background = (selectedButton === idx) ? 'linear-gradient(90deg, #e6f2ff 0%, #f8faff 100%)' : '#f8f8f8'; };
                btn.onclick = function() {
                    selectedButton = idx;
                    updateEditor();
                };
                typeRow.appendChild(btn);
            });
            smsEditorPanel.appendChild(typeRow);

            // Editor wrap and message textarea
            let editorWrap = document.createElement('div');
            editorWrap.className = 'editor-wrap';
            editorWrap.style.display = 'flex';
            editorWrap.style.flexDirection = 'column';
            editorWrap.style.gap = '8px';
            editorWrap.style.marginBottom = '8px';
            smsEditorPanel.appendChild(editorWrap);

            // Flex row for Close and Reset buttons
            let btnRow = document.createElement('div');
            btnRow.style.display = 'flex';
            btnRow.style.justifyContent = 'space-between';
            btnRow.style.alignItems = 'center';
            btnRow.style.marginTop = '14px';

            // Close button
            let closeBtn = document.createElement('button');
            closeBtn.textContent = 'Close';
            closeBtn.style.padding = '4px 14px';
            closeBtn.style.borderRadius = '6px';
            closeBtn.style.border = '1px solid #0077cc';
            closeBtn.style.background = '#ffe6e6';
            closeBtn.style.cursor = 'pointer';
            closeBtn.onclick = function() {
                smsEditorPanel.style.display = 'none';
            };

            // Reset button
            let resetBtn = document.createElement('button');
            resetBtn.textContent = 'Reset All Messages';
            resetBtn.style.padding = '4px 14px';
            resetBtn.style.borderRadius = '6px';
            resetBtn.style.border = '1px solid #0077cc';
            resetBtn.style.background = '#e6f2ff';
            resetBtn.style.cursor = 'pointer';
            resetBtn.onclick = function() {
                if (confirm('Are you sure you want to reset all SMS message templates to their default values? This cannot be undone.')) {
                    // Reset all template keys to default
                    let resetTemplates = {};
                    // Standard and auction days
                    resetTemplates['standard_daysMin'] = 29;
                    resetTemplates['standard_daysMax'] = 59;
                    resetTemplates['auction_daysMin'] = 60;
                    resetTemplates['auction_daysMax'] = 999;
                    // Message keys
                    resetTemplates['stdP1'] = 'This is <employeename> from the U-Haul. I’m reaching out because I really want to help you reclaim your belongings, and you have options. The next step is your unit being listed for auction, and I REALLY don’t want that to happen. What can I do to help? Are you interested in your belongings?';
                    resetTemplates['stdP2'] = 'Hi <firstname>, this is <employeename> from U-Haul. Just checking in to see if you need anything or have questions about your storage unit.';
                    resetTemplates['stdA1'] = 'Hey <altfirstname>, my name is <employeename> and I’m the Storage Manager for U-Haul in this area. I’m reaching out because <firstname> has listed you as their emergency contact for their storage unit. We really need to get in contact with them regarding their belongings. Do you have a good number for us to get ahold of them at? Thank you so much!';
                    resetTemplates['stdA2'] = 'Hi <altfirstname>, this is <employeename> from U-Haul. Just checking in to see if you need anything or have questions about <firstname>\'s storage unit.';
                    resetTemplates['aucP1'] = 'Hey <firstname>, my name is <employeename> and I’m the Storage Manager for U-Haul in this area. We are desperately trying to get in contact with our customer regarding their belongings. Is this by chance your storage unit? Please let me know either way and thank you so much for your help!';
                    resetTemplates['aucP2'] = 'Hi <firstname>, this is <employeename> from U-Haul. Just checking in to see if you need anything or have questions about your storage unit.';
                    resetTemplates['aucA1'] = 'Hey <altfirstname>, my name is <employeename> and I’m the Storage Manager for U-Haul in this area. I’m reaching out because <firstname> has listed you as their emergency contact for their storage unit. We really need to get in contact with them regarding their belongings. Do you have a good number for us to get ahold of them at? Thank you so much!';
                    resetTemplates['aucA2'] = 'Hi <altfirstname>, this is <employeename> from U-Haul. Just checking in to see if you need anything or have questions about <firstname>\'s storage unit.';
                    saveTemplates(resetTemplates);
                    alert('All SMS message templates have been reset to default.');
                    showEditor(); // Refresh editor UI
                }
            };

            btnRow.appendChild(closeBtn);
            btnRow.appendChild(resetBtn);
            smsEditorPanel.appendChild(btnRow);

            // Template keys
            const templateKeys = {
                standard: ['stdP1', 'stdP2', 'stdA1', 'stdA2'],
                auction: ['aucP1', 'aucP2', 'aucA1', 'aucA2']
            };
            // For custom ranges, use keys: custom_{idx}_P1, custom_{idx}_P2, custom_{idx}_A1, custom_{idx}_A2

            // Update editor UI
            function updateEditor() {
                Array.from(typeRow.children).forEach((btn, idx) => {
                    btn.style.background = (selectedButton === idx) ? '#e6f2ff' : '#f8f8f8';
                });
                // Remove all children safely to avoid TrustedHTML CSP error
                while (editorWrap.firstChild) {
                    editorWrap.removeChild(editorWrap.firstChild);
                }
                // Days late config for selected range
                let daysRow = document.createElement('div');
                daysRow.style.display = 'flex';
                daysRow.style.gap = '8px';
                daysRow.style.alignItems = 'center';
                let minLabel = document.createElement('label');
                minLabel.innerText = 'Min days late:';
                let minInput = document.createElement('input');
                minInput.type = 'number';
                let maxLabel = document.createElement('label');
                maxLabel.innerText = 'Max days late:';
                let maxInput = document.createElement('input');
                maxInput.type = 'number';
                minInput.style.width = '60px';
                maxInput.style.width = '60px';
                if (selectedDropdown === 'standard') {
                    // Always load from storage for persistence
                    let storedTemplates = {};
                    try { storedTemplates = JSON.parse(GM_getValue('wss_msg_templates', '{}')); } catch (e) { storedTemplates = {}; }
                    minInput.value = storedTemplates['standard_daysMin'] !== undefined ? storedTemplates['standard_daysMin'] : 29;
                    maxInput.value = storedTemplates['standard_daysMax'] !== undefined ? storedTemplates['standard_daysMax'] : 59;
                } else if (selectedDropdown === 'auction') {
                    let storedTemplates = {};
                    try { storedTemplates = JSON.parse(GM_getValue('wss_msg_templates', '{}')); } catch (e) { storedTemplates = {}; }
                    minInput.value = storedTemplates['auction_daysMin'] !== undefined ? storedTemplates['auction_daysMin'] : 60;
                    maxInput.value = storedTemplates['auction_daysMax'] !== undefined ? storedTemplates['auction_daysMax'] : 999;
                } else if (selectedDropdown.startsWith('custom_')) {
                    let idx = parseInt(selectedDropdown.split('_')[1], 10);
                    let r = customRanges[idx];
                    minInput.value = r ? r.daysMin : '';
                    maxInput.value = r ? r.daysMax : '';
                }
                daysRow.appendChild(minLabel);
                daysRow.appendChild(minInput);
                daysRow.appendChild(maxLabel);
                daysRow.appendChild(maxInput);
                editorWrap.appendChild(daysRow);

                // Placeholders row
                let phRow = document.createElement('div');
                phRow.style.display = 'flex';
                phRow.style.gap = '6px';
                phRow.style.flexWrap = 'wrap';
                placeholders.forEach(ph => {
                    let btn = document.createElement('button');
                    btn.type = 'button';
                    btn.innerText = ph.label;
                    btn.title = ph.desc;
                    btn.style.fontSize = '11px';
                    btn.style.padding = '1px 5px';
                    btn.style.borderRadius = '3px';
                    btn.style.border = '1px solid #aaa';
                    btn.style.background = '#f5f5f5';
                    btn.style.cursor = 'pointer';
                    btn.onclick = function() {
                        msgEditor.focus();
                        const start = msgEditor.selectionStart;
                        const end = msgEditor.selectionEnd;
                        const before = msgEditor.value.substring(0, start);
                        const after = msgEditor.value.substring(end);
                        msgEditor.value = before + ph.label + after;
                        msgEditor.selectionStart = msgEditor.selectionEnd = start + ph.label.length;
                    };
                    phRow.appendChild(btn);
                });
                editorWrap.appendChild(phRow);

                // Message textarea
                let msgEditor = document.createElement('textarea');
                msgEditor.style.width = '100%';
                msgEditor.style.height = '120px';
                msgEditor.style.fontSize = '15px';
                msgEditor.style.borderRadius = '6px';
                msgEditor.style.border = '1px solid #aaa';
                msgEditor.style.padding = '8px';
                // Show correct template from GM storage
                let key;
                if (selectedDropdown === 'standard' || selectedDropdown === 'auction') {
                    key = templateKeys[selectedDropdown][selectedButton];
                } else if (selectedDropdown.startsWith('custom_')) {
                    key = `${selectedDropdown}_${['P1','P2','A1','A2'][selectedButton]}`;
                }
                let storedTemplates = {};
                try { storedTemplates = JSON.parse(GM_getValue('wss_msg_templates', '{}')); } catch (e) { storedTemplates = {}; }
                msgEditor.value = storedTemplates[key] !== undefined ? storedTemplates[key] : '';
                editorWrap.appendChild(msgEditor);

                // Save button
                let saveBtn = document.createElement('button');
                saveBtn.className = 'uh-copy-btn';
                saveBtn.innerText = 'Save';
                saveBtn.style.marginTop = '8px';
                saveBtn.onclick = function() {
                    // Load latest from GM storage to avoid overwriting other templates
                    let latestTemplates = {};
                    try { latestTemplates = JSON.parse(GM_getValue('wss_msg_templates', '{}')); } catch (e) { latestTemplates = {}; }
                    // Save days late for selected range
                    if (selectedDropdown === 'standard' || selectedDropdown === 'auction') {
                        latestTemplates[selectedDropdown + '_daysMin'] = parseInt(minInput.value, 10);
                        latestTemplates[selectedDropdown + '_daysMax'] = parseInt(maxInput.value, 10);
                    } else if (selectedDropdown.startsWith('custom_')) {
                        let idx = parseInt(selectedDropdown.split('_')[1], 10);
                        if (customRanges[idx]) {
                            customRanges[idx].daysMin = parseInt(minInput.value, 10);
                            customRanges[idx].daysMax = parseInt(maxInput.value, 10);
                            saveCustomRanges(customRanges);
                        }
                    }
                    // Save message for selected button and range
                    latestTemplates[key] = msgEditor.value;
                    saveTemplates(latestTemplates);
                    // Update dropdown labels after save
                    stdOpt.textContent = `Standard (days ${latestTemplates['standard_daysMin'] || 29}-${latestTemplates['standard_daysMax'] || 59})`;
                    auctionOpt.textContent = `Auction (${latestTemplates['auction_daysMin'] || 60}+ Days)`;
                    saveBtn.innerText = 'Saved!';
                    setTimeout(() => { saveBtn.innerText = 'Save'; }, 1200);
                };
                editorWrap.appendChild(saveBtn);
            }

            templateDropdown.onchange = function() {
                selectedDropdown = this.value;
                updateEditor();
            };

            updateEditor();
        }

        smsEditorBtn.onclick = function() {
            try {
                console.log('VoiceWSS: SMS Editor button clicked');
                // Always append the editor panel to document.body if not present
                if (!document.body.contains(smsEditorPanel)) {
                    document.body.appendChild(smsEditorPanel);
                }
                // Render the editor UI BEFORE making it visible
                showEditor();
                smsEditorPanel.style.zIndex = '100000';
                smsEditorPanel.style.display = 'block';
                smsEditorPanel.style.visibility = 'visible';
                smsEditorPanel.style.opacity = '1';
            } catch (e) {
                console.error('VoiceWSS: Error showing SMS Editor panel', e);
                alert('Error showing SMS Editor panel: ' + e.message);
            }
        };
    // Add SMS Editor button to dropdown container
    container.appendChild(smsEditorBtn);
        // --- End SMS Editor UI ---
        if (window.location.hostname.includes('messages.google.com')) {
            wrapper = document.createElement('div');
            wrapper.style.display = 'flex';
            wrapper.style.justifyContent = 'center';
            wrapper.style.alignItems = 'flex-start';
            wrapper.style.width = 'auto';
            wrapper.style.flex = '0 0 auto';
            wrapper.style.background = 'transparent';
            wrapper.style.position = 'fixed';
            wrapper.style.top = '8px';
            wrapper.style.left = '50%';
            wrapper.style.transform = 'translateX(-50%)';
            wrapper.style.zIndex = '99999';
            wrapper.appendChild(container);
        } else {
            wrapper = null;
        }

        // --- Three-row layout ---
        // Row 1: Your Name
    let row1 = document.createElement('div');
    row1.style.display = 'flex';
    row1.style.justifyContent = 'center';
    row1.style.alignItems = 'center';
    row1.style.gap = '14px';
    row1.style.marginBottom = '10px';
    row1.style.width = '100%';

    // Move icon (dragHandle), name input, refresh button
    dragHandle.style.position = 'static';
    dragHandle.style.margin = '0 0 0 0';
    dragHandle.style.order = '0';

    let empInput = document.createElement('input');
    empInput.type = 'text';
    empInput.placeholder = 'Enter Your Name';
    empInput.style.margin = '0';
    empInput.style.padding = '4px 10px';
    empInput.style.borderRadius = '7px';
    empInput.style.border = '1px solid #0077cc';
    empInput.style.width = '160px';
    empInput.style.background = 'linear-gradient(90deg, #e6f2ff 0%, #f8faff 100%)';
    empInput.style.fontSize = '14px';
    empInput.style.fontWeight = '500';
    empInput.style.boxShadow = '0 1px 4px rgba(0,0,0,0.08)';
    empInput.style.transition = 'box-shadow 0.2s';
    empInput.onfocus = function() { empInput.style.boxShadow = '0 2px 8px rgba(0,0,0,0.16)'; };
    empInput.onblur = function() { empInput.style.boxShadow = '0 1px 4px rgba(0,0,0,0.08)'; };
    empInput.style.order = '1';

    refreshBtn.style.order = '2';
    refreshBtn.style.margin = '0';

    row1.appendChild(dragHandle);
    row1.appendChild(empInput);
    row1.appendChild(refreshBtn);
    container.appendChild(row1);
        // restore previously saved employee name (persist across refresh)
        try {
            const savedEmp = GM_getValue('wss_emp_name', '');
            if (savedEmp) empInput.value = savedEmp;
        } catch (e) { /* ignore if GM_getValue unavailable */ }
        // save on input so the value persists across reloads
        empInput.addEventListener('input', function() {
            try { GM_setValue('wss_emp_name', empInput.value || ''); } catch (e) { /* ignore */ }
        });
    // Row 2: Contact Dropdown
    let row2 = document.createElement('div');
    row2.style.display = 'flex';
    row2.style.justifyContent = 'center';
    row2.style.alignItems = 'center';
    row2.style.marginBottom = '10px';

        // Unified dropdown for main and alternates
    let select = document.createElement('select');
    select.style.width = '270px';
    select.style.padding = '4px 10px';
    select.style.borderRadius = '7px';
    select.style.border = '1px solid #0077cc';
    select.style.boxShadow = '0 1px 4px rgba(0,0,0,0.08)';
    select.style.fontSize = '14px';
    select.style.fontWeight = '500';
    select.style.background = 'linear-gradient(90deg, #f8faff 0%, #e6f2ff 100%)';
    select.style.marginBottom = '0px';
    select.style.transition = 'box-shadow 0.2s';
    select.onfocus = function() { select.style.boxShadow = '0 2px 8px rgba(0,0,0,0.16)'; };
    select.onblur = function() { select.style.boxShadow = '0 1px 4px rgba(0,0,0,0.08)'; };
    row2.appendChild(select);
    container.appendChild(row2);
    let defaultOpt = document.createElement('option');
    defaultOpt.textContent = 'Select Contact';
    defaultOpt.value = '';
    defaultOpt.style.background = '#e6f2ff';
    defaultOpt.style.color = '#003366';
    select.appendChild(defaultOpt);

        // Build dropdown options: main contact, then alternates indented below
        if (Array.isArray(contacts) && contacts.length > 0) {
            // Color palette for grouping
            const palette = ['#e6f2ff', '#ffe6e6', '#e6ffe6', '#fffbe6', '#e6e6ff', '#f8e6ff'];
            // Load messaged contacts from storage
            let messaged = {};
            try {
                messaged = JSON.parse(GM_getValue('wss_messaged', '{}'));
            } catch (e) { messaged = {}; }
            contacts.forEach((c, i) => {
                let color = palette[i % palette.length];
                let daysLateStr = c.daysLate !== undefined && c.daysLate !== null ? `, ${c.daysLate} days late` : '';
                let displayName = c.name.replace(/\s+(N|AUTOPAY)\s*$/i, '').trim();
                let mainKey = `${displayName}|${c.phone}`;
                // Two alternating readable text colors
                const textColors = ['#003366', '#006400'];
                let textColor = messaged[mainKey] ? '#888' : textColors[i % textColors.length];
                let option = document.createElement('option');
                option.value = `main-${i}`;
                option.textContent = `${displayName} (${c.phone}${daysLateStr})` + (messaged[mainKey] ? ' ✔' : '');
                option.style.background = color;
                option.style.color = textColor;
                option.style.padding = '6px 10px';
                option.style.borderRadius = '8px';
                option.style.marginBottom = '2px';
                option.style.fontSize = '15px';
                option.style.fontWeight = messaged[mainKey] ? '400' : '500';
                select.appendChild(option);
                // Add alternates directly below
                if (Array.isArray(c.alternates) && c.alternates.length > 0) {
                    c.alternates.forEach((alt, j) => {
                        let altOpt = document.createElement('option');
                        altOpt.value = `alt-${i}-${j}`;
                        let altDaysLateStr = alt.daysLate !== undefined && alt.daysLate !== null ? `, ${alt.daysLate} days late` : '';
                        let altKey = `${alt.name}|${alt.phone}`;
                        altOpt.textContent = `⮤ Alt Contact: (${alt.phone}${altDaysLateStr})` + (messaged[altKey] ? ' ✔' : '');
                        altOpt.style.background = color;
                        altOpt.style.color = messaged[altKey] ? '#888' : textColor;
                        altOpt.style.padding = '6px 10px';
                        altOpt.style.borderRadius = '8px';
                        altOpt.style.marginBottom = '2px';
                        altOpt.style.fontSize = '15px';
                        altOpt.style.fontWeight = messaged[altKey] ? '400' : '500';
                        select.appendChild(altOpt);
                    });
                }
            });
        } else {
            let emptyOpt = document.createElement('option');
            emptyOpt.textContent = 'No contacts found';
            emptyOpt.value = '';
            select.appendChild(emptyOpt);
        }

        // expose a refresh helper on the select so other code (and buttons) can re-sync option markers
        select.refreshDropdownMarkers = function() {
            let messaged = {};
            try { messaged = JSON.parse(GM_getValue('wss_messaged', '{}')); } catch (e) { messaged = {}; }
            const palette = ['#e6f2ff', '#ffe6e6', '#e6ffe6', '#fffbe6', '#e6e6ff', '#f8e6ff'];
            const textColors = ['#003366', '#006400'];
            for (let k = 0; k < select.options.length; k++) {
                const opt = select.options[k];
                if (!opt || !opt.value) continue;
                if (opt.value === '') continue;
                if (opt.value.startsWith('main-')) {
                    const mainIdx = parseInt(opt.value.split('-')[1], 10);
                    const c = contacts[mainIdx];
                    if (!c) continue;
                    const displayName = c.name.replace(/\s+(N|AUTOPAY)\s*$/i, '').trim();
                    const daysLateStr = c.daysLate !== undefined && c.daysLate !== null ? `, ${c.daysLate} days late` : '';
                    const key = `${displayName}|${c.phone}`;
                    opt.textContent = `${displayName} (${c.phone}${daysLateStr})` + (messaged[key] ? ' ✔' : '');
                    opt.style.background = palette[mainIdx % palette.length];
                    opt.style.color = messaged[key] ? '#888' : (textColors[mainIdx % textColors.length] || '#003366');
                } else if (opt.value.startsWith('alt-')) {
                    const parts = opt.value.split('-');
                    const mi = parseInt(parts[1], 10);
                    const aj = parseInt(parts[2], 10);
                    const parent = contacts[mi];
                    if (!parent || !Array.isArray(parent.alternates) || !parent.alternates[aj]) continue;
                    const alt = parent.alternates[aj];
                    const altDaysLateStr = alt.daysLate !== undefined && alt.daysLate !== null ? `, ${alt.daysLate} days late` : '';
                    const key = `${alt.name}|${alt.phone}`;
                    opt.textContent = `⮤ Alt Contact: (${alt.phone}${altDaysLateStr})` + (messaged[key] ? ' ✔' : '');
                    opt.style.background = palette[mi % palette.length];
                    opt.style.color = messaged[key] ? '#888' : (textColors[mi % textColors.length] || '#003366');
                }
            }
        };

        // Expose convenient functions on window for console use (page context cannot call GM_* directly)
        try {
            window.wssClearMessaged = async function() {
                try {
                    // Use GM_setValue from the userscript context to clear stored messaged map
                    await GM_setValue('wss_messaged', JSON.stringify({}));
                } catch (e) {
                    console.error('VoiceWSS: failed to clear wss_messaged', e);
                }
                try { if (select && typeof select.refreshDropdownMarkers === 'function') select.refreshDropdownMarkers(); } catch (e) { /* ignore */ }
                console.log('VoiceWSS: wss_messaged cleared and dropdown refreshed');
            };
            window.wssGetMessaged = function() {
                try { return JSON.parse(GM_getValue('wss_messaged', '{}')); } catch (e) { return {}; }
            };
            window.wssRefresh = function() { try { if (select && typeof select.refreshDropdownMarkers === 'function') select.refreshDropdownMarkers(); } catch (e) {} };
            // expose the select element for easy console inspection
            window.wssSelect = select;
        } catch (e) {
            /* in case window isn't writable or GM_* not available yet */
        }

        // Message templates
        function getFirstName(fullName) {
            let parts = fullName.split(',');
            let rawFirst = '';
            if (parts.length > 1) {
                let candidates = parts[1].trim().split(/\s+/);
                for (let candidate of candidates) {
                    if (/^(N|AUTOPAY)$/i.test(candidate)) continue;
                    if (/^[A-Za-z]{2,}$/.test(candidate)) {
                        rawFirst = candidate;
                        break;
                    }
                }
            } else {
                let candidates = fullName.trim().split(/\s+/);
                for (let candidate of candidates) {
                    if (/^(N|AUTOPAY)$/i.test(candidate)) continue;
                    if (/^[A-Za-z]{2,}$/.test(candidate)) {
                        rawFirst = candidate;
                        break;
                    }
                }
            }
            if (rawFirst.length > 0) {
                return rawFirst.charAt(0).toUpperCase() + rawFirst.slice(1).toLowerCase();
            }
            return '';
        }
        function getMessage(contact, empName) {
            let firstName = contact.name ? getFirstName(contact.name) : '';
            let daysLate = contact.daysLate;
            let stored = {};
            try { stored = JSON.parse(GM_getValue('wss_msg_templates', '{}')); } catch (e) { stored = {}; }
            // Load custom ranges from storage
            let customRanges = [];
            try { customRanges = JSON.parse(GM_getValue('wss_custom_ranges', '[]')); } catch (e) { customRanges = []; }
            function fillPlaceholders(str) {
                let altFirst = contact.name ? getFirstName(contact.name) : '';
                let mainFirst = contact.mainName ? getFirstName(contact.mainName) : '';
                return str.replace(/<employeename>/g, empName)
                          .replace(/<firstname>/g, firstName)
                          .replace(/<altfirstname>/g, altFirst)
                          .replace(/<mainfirstname>/g, mainFirst);
            }
            // Use getTemplateKey logic for both built-in and custom ranges
            function getSelectedType() {
                let radios = document.getElementsByName('wss-msg-type-radio');
                let selected = 0; // 0: SMS 1, 1: SMS 2
                for (let r of radios) { if (r.checked) selected = (r.value === 'Alternate') ? 1 : 0; }
                return selected;
            }
            let smsType = getSelectedType();
            // Check custom ranges first
            for (let idx = 0; idx < customRanges.length; idx++) {
                let r = customRanges[idx];
                if (daysLate >= r.daysMin && daysLate <= r.daysMax) {
                    // Use same logic as getTemplateKey
                    let key;
                    if (contact.isAlternate) {
                        key = smsType === 0 ? `custom_${idx}_A1` : `custom_${idx}_A2`;
                    } else {
                        key = smsType === 0 ? `custom_${idx}_P1` : `custom_${idx}_P2`;
                    }
                    let msg = stored[key];
                    if (msg) return fillPlaceholders(msg);
                }
            }
            // Built-in ranges
            const mainMin = stored['standard_daysMin'] !== undefined ? stored['standard_daysMin'] : 29;
            const mainMax = stored['standard_daysMax'] !== undefined ? stored['standard_daysMax'] : 59;
            const auctionMin = stored['auction_daysMin'] !== undefined ? stored['auction_daysMin'] : 60;
            const auctionMax = stored['auction_daysMax'] !== undefined ? stored['auction_daysMax'] : 999;
            let key;
            if (contact.isAlternate) {
                if (daysLate >= mainMin && daysLate <= mainMax) {
                    key = smsType === 0 ? 'stdA1' : 'stdA2';
                } else if (daysLate >= auctionMin && daysLate <= auctionMax) {
                    key = smsType === 0 ? 'aucA1' : 'aucA2';
                } else {
                    key = smsType === 0 ? 'stdA1' : 'stdA2';
                }
            } else {
                if (daysLate >= mainMin && daysLate <= mainMax) {
                    key = smsType === 0 ? 'stdP1' : 'stdP2';
                } else if (daysLate >= auctionMin && daysLate <= auctionMax) {
                    key = smsType === 0 ? 'aucP1' : 'aucP2';
                } else {
                    key = smsType === 0 ? 'stdP1' : 'stdP2';
                }
            }
            let msg = stored[key];
            if (msg) return fillPlaceholders(msg);
            // Fallback
            return fillPlaceholders('Hi <firstname>, this is <employeename> from U-Haul. Just checking in to see if you need anything or have questions about your storage unit.');
        }

        select.addEventListener('change', function() {
            let val = this.value;
            let contact = null;
            let isAlt = false;
            let mainIdx = null, altIdx = null;
            if (/^main-(\d+)$/.test(val)) {
                mainIdx = parseInt(val.split('-')[1], 10);
                contact = contacts[mainIdx];
            } else if (/^alt-(\d+)-(\d+)$/.test(val)) {
                let parts = val.split('-');
                mainIdx = parseInt(parts[1], 10);
                altIdx = parseInt(parts[2], 10);
                contact = Object.assign({}, contacts[mainIdx].alternates[altIdx]);
                contact.isAlternate = true;
                contact.mainName = contacts[mainIdx].name;
            }
            (async function markSelectedAsMessaged() {
                try {
                    if (!contact) return;
                    let key = contact.isAlternate ? `${contact.name}|${contact.phone}` : `${contact.name.replace(/\s+(N|AUTOPAY)\s*$/i, '').trim()}|${contact.phone}`;
                    let messaged = {};
                    try { messaged = JSON.parse(GM_getValue('wss_messaged', '{}')); } catch (e) { messaged = {}; }
                    messaged[key] = true;
                    try { await GM_setValue('wss_messaged', JSON.stringify(messaged)); } catch (e) { try { GM_setValue('wss_messaged', JSON.stringify(messaged)); } catch (ee) {} }
                    try {
                        const opt = select.querySelector(`option[value="${val}"]`);
                        if (opt) {
                            if (!/✔$/.test(opt.textContent)) opt.textContent = opt.textContent + ' ✔';
                            opt.style.color = '#888';
                        }
                    } catch (e) { /* ignore */ }
                    try { if (select && typeof select.refreshDropdownMarkers === 'function') select.refreshDropdownMarkers(); } catch (e) { /* ignore */ }
                } catch (e) { console.error('VoiceWSS: markSelectedAsMessaged error', e); }
            })();
            // Autofill logic for both platforms
            function getSelectedType() {
                let radios = document.getElementsByName('wss-msg-type-radio');
                let selected = 0; // 0: SMS 1, 1: SMS 2
                for (let r of radios) { if (r.checked) selected = (r.value === 'Alternate') ? 1 : 0; }
                return selected;
            }
            function getTemplateKey(contact, daysLate, smsType) {
                // smsType: 0 (SMS 1), 1 (SMS 2)
                let stored = {};
                try { stored = JSON.parse(GM_getValue('wss_msg_templates', '{}')); } catch (e) { stored = {}; }
                let stdMin = stored['standard_daysMin'] !== undefined ? stored['standard_daysMin'] : 29;
                let stdMax = stored['standard_daysMax'] !== undefined ? stored['standard_daysMax'] : 59;
                let aucMin = stored['auction_daysMin'] !== undefined ? stored['auction_daysMin'] : 60;
                let aucMax = stored['auction_daysMax'] !== undefined ? stored['auction_daysMax'] : 999;
                let range = (daysLate >= stdMin && daysLate <= stdMax) ? 'standard' : (daysLate >= aucMin && daysLate <= aucMax ? 'auction' : 'standard');
                if (contact.isAlternate) {
                    // Alternate contact
                    return range === 'standard' ? (smsType === 0 ? 'stdA1' : 'stdA2') : (smsType === 0 ? 'aucA1' : 'aucA2');
                } else {
                    // Primary contact
                    return range === 'standard' ? (smsType === 0 ? 'stdP1' : 'stdP2') : (smsType === 0 ? 'aucP1' : 'aucP2');
                }
            }
            function fillPlaceholders(str, contact, empName) {
                // <firstname>: always main contact's first name
                // <altfirstname>: always alternate contact's first name
                let mainFirst = '';
                let altFirst = '';
                if (contact.isAlternate) {
                    mainFirst = contact.mainName ? getFirstName(contact.mainName) : '';
                    altFirst = contact.name ? getFirstName(contact.name) : '';
                } else {
                    mainFirst = contact.name ? getFirstName(contact.name) : '';
                    if (Array.isArray(contact.alternates) && contact.alternates.length > 0) {
                        altFirst = getFirstName(contact.alternates[0].name);
                    } else {
                        altFirst = '';
                    }
                }
                return str.replace(/<employeename>/g, empName)
                          .replace(/<firstname>/g, mainFirst)
                          .replace(/<altfirstname>/g, altFirst)
                          .replace(/<mainfirstname>/g, mainFirst);
            }
            function getPopulatedMessage(contact, empName) {
                let daysLate = contact.daysLate;
                    let smsType = getSelectedType();
                    let stored = {};
                    try { stored = JSON.parse(GM_getValue('wss_msg_templates', '{}')); } catch (e) { stored = {}; }
                    // Load custom ranges from storage
                    let customRanges = [];
                    try { customRanges = JSON.parse(GM_getValue('wss_custom_ranges', '[]')); } catch (e) { customRanges = []; }
                    // Check custom ranges first
                    for (let idx = 0; idx < customRanges.length; idx++) {
                        let r = customRanges[idx];
                        if (daysLate >= r.daysMin && daysLate <= r.daysMax) {
                            let key;
                            if (contact.isAlternate) {
                                key = smsType === 0 ? `custom_${idx}_A1` : `custom_${idx}_A2`;
                            } else {
                                key = smsType === 0 ? `custom_${idx}_P1` : `custom_${idx}_P2`;
                            }
                            let msg = stored[key];
                            if (msg) return fillPlaceholders(msg, contact, empName);
                        }
                    }
                    // Built-in ranges
                    let key = getTemplateKey(contact, daysLate, smsType);
                    let msg = stored[key] || '';
                    return fillPlaceholders(msg, contact, empName);
            }
            function updateMessageBox() {
                let empName = empInput.value || '';
                let messageInput = document.querySelector('textarea[placeholder*="message"], textarea[aria-label*="message"], input[placeholder*="message"], input[aria-label*="message"]');
                if (!messageInput) return;
                messageInput.value = getPopulatedMessage(contact, empName);
                messageInput.dispatchEvent(new Event('input', { bubbles: true }));
            }
            // Listen for radio changes to update message
            let radios = document.getElementsByName('wss-msg-type-radio');
            for (let r of radios) r.onchange = updateMessageBox;
            // Autofill logic for both platforms
            if (window.location.hostname.includes('voice.google.com')) {
                try { openComposerAndFill(contact); } catch (e) { /* ignore if helper not ready */ }
                setTimeout(updateMessageBox, 600);
            } else if (window.location.hostname.includes('messages.google.com')) {
                let startChatBtn = Array.from(document.querySelectorAll('div.fab-label')).find(el => el.textContent.trim().toLowerCase() === 'start chat');
                if (startChatBtn) {
                    startChatBtn.click();
                }
                setTimeout(() => {
                    let input = document.querySelector('input[data-e2e-contact-input]');
                    if (input) {
                        input.value = contact.phone;
                        input.dispatchEvent(new Event('input', { bubbles: true }));
                        input.focus();
                        setTimeout(() => {
                            function normalizePhone(str) {
                                return str.replace(/\D/g, '');
                            }
                            let targetPhone = normalizePhone(contact.phone);
                            let spans = Array.from(document.querySelectorAll('span.anon-contact-name'));
                            let found = spans.find(span => normalizePhone(span.textContent) === targetPhone);
                            if (found) {
                                found.click();
                            } else {
                                const enterEvent = new KeyboardEvent('keydown', {
                                    key: 'Enter',
                                    code: 'Enter',
                                    bubbles: true,
                                    cancelable: true,
                                });
                                input.dispatchEvent(enterEvent);
                            }
                        }, 300);
                        let msgAttempts = 0;
                        const msgMaxAttempts = 50;
                        const msgPoll = setInterval(() => {
                            let empName = empInput.value || '';
                            let msgEl = document.querySelector('textarea[data-e2e-message-input-box]');
                            if (msgEl) {
                                clearInterval(msgPoll);
                                let messageText = getPopulatedMessage(contact, empName);
                                msgEl.value = messageText;
                                msgEl.dispatchEvent(new Event('input', { bubbles: true }));
                                msgEl.dispatchEvent(new Event('change', { bubbles: true }));
                                msgEl.focus();
                            } else if (msgAttempts++ > msgMaxAttempts) {
                                clearInterval(msgPoll);
                                console.log('VoiceWSS: Timed out waiting for message box on Google Messages.');
                            }
                        }, 100);
                    }
                }, 500);
            }
        });

        // Helper: click Send New Message and populate recipient field with name/phone
        function openComposerAndFill(contact) {
            if (!contact) return;
            try {
                let sendBtn = Array.from(document.querySelectorAll('div.gmat-subhead-2.grey-900')).find(d => d.textContent && d.textContent.trim().toLowerCase().includes('send new message'));
                if (sendBtn) {
                    sendBtn.click();
                }
            } catch (e) { /* ignore */ }
            let attempts = 0;
            const maxAttempts = 40;
            const poll = setInterval(() => {
                attempts++;
                const selectors = [
                    'input[placeholder="Type a name or phone number"]',
                    'input[id^="mat-mdc-chip-list-input-"]',
                    'input[aria-label="To"]',
                    'input[aria-label*="Search"]',
                    'input[placeholder*="Search"]',
                    'input[placeholder*="To"]',
                    'input[role="combobox"]',
                    'input[type="text"]'
                ];
                let input = null;
                for (let s of selectors) {
                    let el = document.querySelector(s);
                    if (el && el.offsetParent !== null) { input = el; break; }
                }
                if (!input) {
                    let editable = document.querySelector('div[contenteditable="true"][role="combobox"]');
                    if (editable) input = editable;
                }
                if (input) {
                    try {
                        const val = contact.phone || (contact.name || '');
                        if (input.tagName.toLowerCase() === 'input') {
                            input.focus();
                            input.value = val;
                            input.dispatchEvent(new Event('input', { bubbles: true }));
                            input.dispatchEvent(new Event('change', { bubbles: true }));
                            input.dispatchEvent(new KeyboardEvent('keydown', { key: 'a' }));
                            input.dispatchEvent(new KeyboardEvent('keyup', { key: 'a' }));
                        } else {
                            input.focus();
                            input.textContent = val;
                            input.dispatchEvent(new InputEvent('input', { bubbles: true }));
                        }
                        let sendToAttempts = 0;
                        const sendToMax = 30;
                        const sendToPoll = setInterval(() => {
                            sendToAttempts++;
                            let sendCandidate = document.querySelector('div.send-to-button, div[id^="send_to_button-"]');
                            if (sendCandidate && sendCandidate.offsetParent !== null) {
                                try { sendCandidate.click(); } catch (e) { /* ignore click errors */ }
                                clearInterval(sendToPoll);
                            } else if (sendToAttempts >= sendToMax) {
                                clearInterval(sendToPoll);
                            }
                        }, 120);
                        const empName = (typeof empInput !== 'undefined' && empInput) ? empInput.value || '' : '';
                        let msgAttempts = 0;
                        const msgMax = 40;
                        const msgPoll = setInterval(() => {
                            msgAttempts++;
                            const msgSelectors = [
                                'textarea[placeholder*="message"]',
                                'textarea[aria-label*="message"]',
                                'div[aria-label*="Message"]',
                                'div[role="textbox"][contenteditable="true"]',
                                'input[aria-label*="message"]',
                                'input[placeholder*="message"]'
                            ];
                            let msgEl = null;
                            for (let s of msgSelectors) {
                                const el = document.querySelector(s);
                                if (el && el.offsetParent !== null) { msgEl = el; break; }
                            }
                            if (msgEl) {
                                try {
                                    const messageText = (typeof getMessage === 'function') ? getMessage(contact, empName) : '';
                                    if (msgEl.tagName.toLowerCase() === 'textarea' || msgEl.tagName.toLowerCase() === 'input') {
                                        msgEl.focus();
                                        msgEl.value = messageText;
                                        msgEl.dispatchEvent(new Event('input', { bubbles: true }));
                                        msgEl.dispatchEvent(new Event('change', { bubbles: true }));
                                    } else {
                                        msgEl.focus();
                                        msgEl.textContent = messageText;
                                        msgEl.dispatchEvent(new InputEvent('input', { bubbles: true }));
                                    }
                                } catch (e) { /* ignore */ }
                                clearInterval(msgPoll);
                            } else if (msgAttempts >= msgMax) {
                                clearInterval(msgPoll);
                            }
                        }, 120);
                    } catch (e) { /* ignore */ }
                    clearInterval(poll);
                } else if (attempts >= maxAttempts) {
                    clearInterval(poll);
                }
            }, 120);
        }

        // Row 3: Message type selector and SMS Editor button
        let row3 = document.createElement('div');
        row3.style.display = 'flex';
        row3.style.justifyContent = 'center';
        row3.style.alignItems = 'center';
        row3.style.marginBottom = '0px';
        row3.style.gap = '18px';
        let msgTypeContainer = document.createElement('div');
        msgTypeContainer.id = 'wss-msg-type-container';
        msgTypeContainer.style.display = 'flex';
        msgTypeContainer.style.alignItems = 'center';
        msgTypeContainer.style.gap = '18px';
        // SMS 1 radio
        let sms1Label = document.createElement('label');
        sms1Label.style.display = 'flex';
        sms1Label.style.alignItems = 'center';
        sms1Label.style.marginRight = '0px';
        let sms1Radio = document.createElement('input');
        sms1Radio.type = 'radio';
        sms1Radio.name = 'wss-msg-type-radio';
        sms1Radio.value = 'Standard';
        sms1Radio.checked = true;
        sms1Radio.style.marginRight = '5px';
        sms1Label.appendChild(sms1Radio);
        let sms1Text = document.createElement('span');
        sms1Text.textContent = 'SMS 1';
        sms1Text.style.whiteSpace = 'nowrap';
        sms1Label.appendChild(sms1Text);
        msgTypeContainer.appendChild(sms1Label);
        // SMS 2 radio
        let sms2Label = document.createElement('label');
        sms2Label.style.display = 'flex';
        sms2Label.style.alignItems = 'center';
        sms2Label.style.marginRight = '0px';
        let sms2Radio = document.createElement('input');
        sms2Radio.type = 'radio';
        sms2Radio.name = 'wss-msg-type-radio';
        sms2Radio.value = 'Alternate';
        sms2Radio.style.marginRight = '5px';
        sms2Label.appendChild(sms2Radio);
        let sms2Text = document.createElement('span');
        sms2Text.textContent = 'SMS 2';
        sms2Text.style.whiteSpace = 'nowrap';
        sms2Label.appendChild(sms2Text);
        msgTypeContainer.appendChild(sms2Label);
        // SMS Editor button
        smsEditorBtn.style.marginLeft = '10px';
        smsEditorBtn.style.marginRight = '0px';
        msgTypeContainer.appendChild(smsEditorBtn);
        row3.appendChild(msgTypeContainer);
        container.appendChild(row3);
        // Prefer main-container for Google Messages
        let mainContainer = document.querySelector('div.main-container[data-e2e-main-container]');
        if (mainContainer) {
            if (wrapper) {
                mainContainer.insertBefore(wrapper, mainContainer.firstChild);
            } else {
                mainContainer.insertBefore(container, mainContainer.firstChild);
            }
        } else {
            let firstContent = document.querySelector('div.content');
            if (firstContent) {
                if (wrapper) {
                    firstContent.insertBefore(wrapper, firstContent.firstChild);
                } else {
                    firstContent.insertBefore(container, firstContent.firstChild);
                }
            } else if (document.body) {
                if (wrapper) {
                    document.body.insertBefore(wrapper, document.body.firstChild);
                } else {
                    document.body.insertBefore(container, document.body.firstChild);
                }
            }
        }
    }

    function shouldShowDropdown() {
        // Only show dropdown on Google Voice or Google Messages main/new/conversation pages
        if (window.location.hostname.includes('voice.google.com')) return true;
        if (window.location.hostname.includes('messages.google.com')) {
            // Show on /web, /web/conversations, /web/conversations/new, /web/conversations/*
            const path = window.location.pathname;
            return (
                path === '/web' ||
                path === '/web/conversations' ||
                path === '/web/conversations/new' ||
                /^\/web\/conversations\/.+/.test(path)
            );
        }
        return false;
    }

    let lastUrl = location.href;
    function checkAndInject() {
        if (shouldShowDropdown()) {
            // Wait for main content container before injecting
            let attempts = 0;
            const maxAttempts = 40;
            const poll = setInterval(() => {
                attempts++;
                let firstContent = document.querySelector('div.content') || document.body;
                if (firstContent) {
                    injectDropdown();
                    clearInterval(poll);
                    // After injection, keep re-injecting for 1 second if dropdown disappears
                    let keepAliveAttempts = 0;
                    const keepAliveMax = 8; // 8 * 125ms = 1 second
                    const keepAlive = setInterval(() => {
                        keepAliveAttempts++;
                        if (shouldShowDropdown() && !document.querySelector('.wss-contact-dropdown')) {
                            injectDropdown();
                        }
                        if (keepAliveAttempts >= keepAliveMax) {
                            clearInterval(keepAlive);
                        }
                    }, 125);
                } else if (attempts >= maxAttempts) {
                    clearInterval(poll);
                }
            }, 120);
        } else {
            // Remove dropdown if present on unsupported page
            let dd = document.querySelector('.wss-contact-dropdown');
            if (dd) dd.remove();
        }
    }

    if (window.location.hostname.includes('voice.google.com') || window.location.hostname.includes('messages.google.com')) {
        checkAndInject();
        // Observe body for DOM changes and reinject if needed
        const observer = new MutationObserver(() => {
            checkAndInject();
        });
        observer.observe(document.body, { childList: true, subtree: true });
        // Also poll for URL changes (SPA navigation)
        setInterval(() => {
            if (location.href !== lastUrl) {
                lastUrl = location.href;
                checkAndInject();
            }
        }, 500);
    }

})();
