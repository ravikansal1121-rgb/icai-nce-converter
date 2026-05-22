document.addEventListener('DOMContentLoaded', function() {
    const cyFile = document.getElementById('cyFile');
    const pyFile = document.getElementById('pyFile');
    const convertBtn = document.getElementById('convertBtn');
    const form = document.getElementById('convertForm');

    function checkReady() {
        convertBtn.disabled = !(cyFile.files.length > 0 && pyFile.files.length > 0);
    }

    function setupDropZone(zoneId, fileInput, infoId, nameId) {
        const zone = document.getElementById(zoneId);
        const info = document.getElementById(infoId);
        const nameEl = document.getElementById(nameId);

        ['dragenter', 'dragover'].forEach(evt => {
            zone.addEventListener(evt, function(e) {
                e.preventDefault();
                zone.classList.add('drag-over');
            });
        });

        ['dragleave', 'drop'].forEach(evt => {
            zone.addEventListener(evt, function(e) {
                e.preventDefault();
                zone.classList.remove('drag-over');
            });
        });

        zone.addEventListener('drop', function(e) {
            const files = e.dataTransfer.files;
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
        const fileInput = document.getElementById(type + 'File');
        const zone = document.getElementById(type + 'Zone');
        const info = document.getElementById(type + 'FileInfo');

        fileInput.value = '';
        zone.querySelector('.upload-content').classList.remove('d-none');
        info.classList.add('d-none');
        checkReady();
    };

    setupDropZone('cyZone', cyFile, 'cyFileInfo', 'cyFileName');
    setupDropZone('pyZone', pyFile, 'pyFileInfo', 'pyFileName');

    form.addEventListener('submit', function(e) {
        e.preventDefault();

        const formData = new FormData(form);

        document.getElementById('progressCard').classList.remove('d-none');
        document.getElementById('resultsCard').classList.add('d-none');
        document.getElementById('errorCard').classList.add('d-none');
        convertBtn.disabled = true;

        fetch('/convert', {
            method: 'POST',
            body: formData
        })
        .then(res => res.json().then(data => ({ ok: res.ok, data })))
        .then(({ ok, data }) => {
            document.getElementById('progressCard').classList.add('d-none');

            if (!ok || data.error) {
                document.getElementById('errorCard').classList.remove('d-none');
                document.getElementById('errorText').textContent = data.error || 'Unknown error';
                convertBtn.disabled = false;
                return;
            }

            document.getElementById('resultsCard').classList.remove('d-none');

            document.getElementById('resEntity').textContent = data.entity_name || '-';
            document.getElementById('resConstitution').textContent = (data.constitution || '-').charAt(0).toUpperCase() + (data.constitution || '').slice(1);
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

            const suppList = data.suppressed_list || [];
            document.getElementById('resSuppressedList').textContent =
                suppList.length > 0 ? suppList.join(', ') : 'None';

            document.getElementById('downloadBtn').href = '/download/' + data.job_id;
            convertBtn.disabled = false;
        })
        .catch(err => {
            document.getElementById('progressCard').classList.add('d-none');
            document.getElementById('errorCard').classList.remove('d-none');
            document.getElementById('errorText').textContent = 'Network error: ' + err.message;
            convertBtn.disabled = false;
        });
    });

    function formatNum(val) {
        if (val === null || val === undefined) return '0';
        const num = parseFloat(val);
        if (isNaN(num)) return '0';
        if (num >= 10000000) {
            return (num / 10000000).toFixed(2) + ' Cr';
        } else if (num >= 100000) {
            return (num / 100000).toFixed(2) + ' L';
        }
        return num.toLocaleString('en-IN', { maximumFractionDigits: 0 });
    }
});
