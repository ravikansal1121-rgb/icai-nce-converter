document.addEventListener('DOMContentLoaded', function() {
    const cyFile = document.getElementById('cyFile');
    const pyFile = document.getElementById('pyFile');
    const convertBtn = document.getElementById('convertBtn');
    const form = document.getElementById('convertForm');
    const singleFileCheck = document.getElementById('singleFileCheck');
    const singleFileNote = document.getElementById('singleFileNote');
    const pyUploadCol = document.getElementById('pyUploadCol');
    const partnershipSection = document.getElementById('partnershipSection');
    const partnerRows = document.getElementById('partnerRows');
    const addPartnerBtn = document.getElementById('addPartnerBtn');

    // ---- Constitution toggle ----
    const constitutionRadios = document.querySelectorAll('input[name="constitution"]');
    const optProprietorship = document.getElementById('optProprietorship');
    const optPartnership = document.getElementById('optPartnership');

    constitutionRadios.forEach(function(radio) {
        radio.addEventListener('change', function() {
            // Update visual selection
            optProprietorship.classList.toggle('selected', this.value === 'proprietorship');
            optPartnership.classList.toggle('selected', this.value === 'partnership');
            // Show/hide partnership section
            if (this.value === 'partnership') {
                partnershipSection.classList.remove('d-none');
            } else {
                partnershipSection.classList.add('d-none');
            }
            checkReady();
        });
    });

    // Click on the whole label card to select radio
    [optProprietorship, optPartnership].forEach(function(opt) {
        opt.addEventListener('click', function() {
            var radio = opt.querySelector('input[type="radio"]');
            radio.checked = true;
            radio.dispatchEvent(new Event('change'));
        });
    });

    // ---- Single file checkbox ----
    singleFileCheck.addEventListener('change', function() {
        if (this.checked) {
            pyUploadCol.classList.add('d-none');
            singleFileNote.classList.remove('d-none');
            // Clear PY file if any
            clearFile('py');
        } else {
            pyUploadCol.classList.remove('d-none');
            singleFileNote.classList.add('d-none');
        }
        checkReady();
    });

    // ---- Partner rows ----
    var partnerCount = 2;

    addPartnerBtn.addEventListener('click', function() {
        partnerCount++;
        var tr = document.createElement('tr');
        tr.setAttribute('data-partner-idx', partnerCount);
        tr.innerHTML =
            '<td class="partner-num">' + partnerCount + '</td>' +
            '<td><input type="text" class="form-control form-control-sm" name="partner_names[]" placeholder="Partner name"></td>' +
            '<td><input type="number" class="form-control form-control-sm" name="partner_shares[]" placeholder="%" min="0" max="100" step="0.01"></td>' +
            '<td><button type="button" class="btn btn-sm btn-outline-danger remove-partner-btn"><i class="bi bi-x-lg"></i></button></td>';
        partnerRows.appendChild(tr);
        renumberPartners();
    });

    // Delegate remove button clicks
    partnerRows.addEventListener('click', function(e) {
        var btn = e.target.closest('.remove-partner-btn');
        if (!btn) return;
        var rows = partnerRows.querySelectorAll('tr');
        if (rows.length <= 2) {
            alert('Minimum 2 partners required for a partnership firm.');
            return;
        }
        btn.closest('tr').remove();
        renumberPartners();
    });

    function renumberPartners() {
        var rows = partnerRows.querySelectorAll('tr');
        rows.forEach(function(row, idx) {
            row.querySelector('.partner-num').textContent = idx + 1;
        });
        partnerCount = rows.length;
    }

    // ---- Client info chevron toggle ----
    var clientInfoBody = document.getElementById('clientInfoBody');
    var clientInfoChevron = document.getElementById('clientInfoChevron');
    if (clientInfoBody) {
        clientInfoBody.addEventListener('show.bs.collapse', function() {
            clientInfoChevron.classList.remove('bi-chevron-down');
            clientInfoChevron.classList.add('bi-chevron-up');
        });
        clientInfoBody.addEventListener('hide.bs.collapse', function() {
            clientInfoChevron.classList.remove('bi-chevron-up');
            clientInfoChevron.classList.add('bi-chevron-down');
        });
    }

    // ---- Readiness check ----
    function checkReady() {
        var hasCy = cyFile.files.length > 0;
        var isSingle = singleFileCheck.checked;
        var hasPy = pyFile.files.length > 0;
        var constitutionSelected = document.querySelector('input[name="constitution"]:checked') !== null;

        if (!constitutionSelected) {
            convertBtn.disabled = true;
            return;
        }

        if (isSingle) {
            convertBtn.disabled = !hasCy;
        } else {
            convertBtn.disabled = !(hasCy && hasPy);
        }
    }

    // ---- File upload / drop zones ----
    function setupDropZone(zoneId, fileInput, infoId, nameId) {
        var zone = document.getElementById(zoneId);
        var info = document.getElementById(infoId);
        var nameEl = document.getElementById(nameId);

        ['dragenter', 'dragover'].forEach(function(evt) {
            zone.addEventListener(evt, function(e) {
                e.preventDefault();
                zone.classList.add('drag-over');
            });
        });

        ['dragleave', 'drop'].forEach(function(evt) {
            zone.addEventListener(evt, function(e) {
                e.preventDefault();
                zone.classList.remove('drag-over');
            });
        });

        zone.addEventListener('drop', function(e) {
            var files = e.dataTransfer.files;
            if (files.length > 0) {
                fileInput.files = files;
                showFileInfo(zone, info, nameEl, files[0].name);
                checkReady();
            }
        });

        fileInput.addEventListener('change', function() {
            if (fileInput.files.length > 0) {
                showFileInfo(zone, info, nameEl, fileInput.files[0].name);
                checkReady();
            }
        });
    }

    function showFileInfo(zone, info, nameEl, filename) {
        zone.querySelector('.upload-content').classList.add('d-none');
        info.classList.remove('d-none');
        nameEl.textContent = filename;
    }

    window.clearFile = function(type) {
        var fileInput = document.getElementById(type + 'File');
        var zone = document.getElementById(type + 'Zone');
        var info = document.getElementById(type + 'FileInfo');

        fileInput.value = '';
        zone.querySelector('.upload-content').classList.remove('d-none');
        info.classList.add('d-none');
        checkReady();
    };

    setupDropZone('cyZone', cyFile, 'cyFileInfo', 'cyFileName');
    setupDropZone('pyZone', pyFile, 'pyFileInfo', 'pyFileName');

    // ---- Form submission ----
    form.addEventListener('submit', function(e) {
        e.preventDefault();

        var formData = new FormData(form);

        // The form already includes all named inputs (constitution, entity_name, pan, gstin,
        // business_nature, business_description, partner_names[], partner_shares[],
        // interest_on_capital_rate, partner_remuneration, single_file, cy_file, py_file)
        // via the standard FormData constructor.

        document.getElementById('progressCard').classList.remove('d-none');
        document.getElementById('resultsCard').classList.add('d-none');
        document.getElementById('errorCard').classList.add('d-none');
        convertBtn.disabled = true;

        fetch('/convert', {
            method: 'POST',
            body: formData
        })
        .then(function(res) {
            return res.json().then(function(data) {
                return { ok: res.ok, data: data };
            });
        })
        .then(function(result) {
            var ok = result.ok;
            var data = result.data;

            document.getElementById('progressCard').classList.add('d-none');

            if (!ok || data.error) {
                document.getElementById('errorCard').classList.remove('d-none');
                document.getElementById('errorText').textContent = data.error || 'Unknown error';
                convertBtn.disabled = false;
                return;
            }

            document.getElementById('resultsCard').classList.remove('d-none');

            document.getElementById('resEntity').textContent = data.entity_name || '-';
            document.getElementById('resConstitution').textContent =
                (data.constitution || '-').charAt(0).toUpperCase() + (data.constitution || '').slice(1);
            document.getElementById('resFyCy').textContent = data.fy_cy || '-';
            document.getElementById('resFyPy').textContent = data.fy_py || '-';

            document.getElementById('resRevCy').textContent = formatNum(data.revenue_cy);
            document.getElementById('resRevPy').textContent = formatNum(data.revenue_py);
            document.getElementById('resNpCy').textContent = formatNum(data.net_profit_cy);
            document.getElementById('resNpPy').textContent = formatNum(data.net_profit_py);
            document.getElementById('resCapCy').textContent = formatNum(data.capital_cy);
            document.getElementById('resCapPy').textContent = formatNum(data.capital_py);

            document.getElementById('resRetained').textContent = data.retained_notes || 0;
            document.getElementById('resSuppressed').textContent = data.suppressed_notes || 0;
            document.getElementById('resNoteRange').textContent = data.note_range || '-';

            var suppList = data.suppressed_list || [];
            document.getElementById('resSuppressedList').textContent =
                suppList.length > 0 ? suppList.join(', ') : 'None';

            document.getElementById('downloadBtn').href = '/download/' + data.job_id;
            convertBtn.disabled = false;
        })
        .catch(function(err) {
            document.getElementById('progressCard').classList.add('d-none');
            document.getElementById('errorCard').classList.remove('d-none');
            document.getElementById('errorText').textContent = 'Network error: ' + err.message;
            convertBtn.disabled = false;
        });
    });

    function formatNum(val) {
        if (val === null || val === undefined) return '0';
        var num = parseFloat(val);
        if (isNaN(num)) return '0';
        if (num >= 10000000) {
            return (num / 10000000).toFixed(2) + ' Cr';
        } else if (num >= 100000) {
            return (num / 100000).toFixed(2) + ' L';
        }
        return num.toLocaleString('en-IN', { maximumFractionDigits: 0 });
    }
});
