with open(r"C:\Users\IT_OFFICE02\Desktop\Chenab_Result_Portal\build_desktop_portal.py", "r", encoding="utf-8") as f:
    text = f.read()

# 1. Enlarge signature CSS even more for prominence
text = text.replace(
    """.sig-img-container {{
            height: 82px;
            display: flex;
            align-items: flex-end;
            justify-content: center;
            margin-bottom: 3px;
        }}
        .sig-img {{
            max-height: 78px;
            max-width: 195px;
            object-fit: contain;
            filter: contrast(1.15) brightness(0.95);
        }}""",
    """.sig-img-container {{
            height: 85px;
            display: flex;
            align-items: flex-end;
            justify-content: center;
            margin-bottom: 3px;
        }}
        .sig-img {{
            max-height: 80px;
            max-width: 200px;
            object-fit: contain;
            filter: contrast(1.15) brightness(0.95);
        }}"""
)

text = text.replace(
    """.sig-img-container {{
                height: 60px !important;
                margin-bottom: 1px !important;
            }}
            .sig-img {{
                max-height: 58px !important;
                max-width: 155px !important;
            }}""",
    """.sig-img-container {{
                height: 65px !important;
                margin-bottom: 1px !important;
            }}
            .sig-img {{
                max-height: 62px !important;
                max-width: 165px !important;
            }}"""
)

# 2. Fix QR code /200 hardcoding
text = text.replace(
    "Marks: ' + st.obtained_marks + '/200 ('",
    "Marks: ' + st.obtained_marks + '/' + st.total_marks + ' ('"
)

# 3. Add classSelect dropdown to modal HTML
old_modal_controls = '''<div class="honor-control-bar">
                    <div class="honor-segmented-switch">
                        <button type="button" class="switch-opt active" id="btnModeClass" onclick="setHonorMode('class')">
                            ? Entire Class 7
                        </button>
                        <button type="button" class="switch-opt" id="btnModeSection" onclick="setHonorMode('section')">
                            ? Section-Wise
                        </button>
                    </div>

                    <div id="sectionDropdownWrap" style="display: none;">
                        <select id="sectionSelect" class="section-select-dropdown" onchange="renderHonorToppers()">
                            <!-- Populated dynamically -->
                        </select>
                    </div>
                </div>'''

new_modal_controls = '''<div class="honor-control-bar">
                    <div class="honor-segmented-switch">
                        <button type="button" class="switch-opt active" id="btnModeSection" onclick="setHonorMode('section')">
                            🏆 Section-Wise
                        </button>
                        <button type="button" class="switch-opt" id="btnModeClass" onclick="setHonorMode('class')">
                            🌐 Entire Class
                        </button>
                    </div>

                    <div id="sectionDropdownWrap">
                        <select id="sectionSelect" class="section-select-dropdown" onchange="renderHonorToppers()">
                            <!-- Populated dynamically -->
                        </select>
                    </div>
                    <div id="classDropdownWrap" style="display: none;">
                        <select id="classSelect" class="section-select-dropdown" onchange="renderHonorToppers()">
                            <!-- Populated dynamically -->
                        </select>
                    </div>
                </div>'''

if old_modal_controls in text:
    text = text.replace(old_modal_controls, new_modal_controls)
    print("Updated modal controls HTML.")
else:
    print("Warning: old_modal_controls not matched directly!")

# 4. Update modal JS functions
old_modal_js = '''function openHonorModal() {{
            const dropdown = document.getElementById('sectionSelect');
            dropdown.innerHTML = '';
            DB.sections.forEach(sec => {{
                const opt = document.createElement('option');
                opt.value = sec;
                opt.textContent = sec;
                dropdown.appendChild(opt);
            }});

            renderHonorToppers();
            document.getElementById('honorModal').style.display = 'flex';
        }}

        function closeHonorModal() {{
            document.getElementById('honorModal').style.display = 'none';
        }}

        function handleBackdropClick(e) {{
            if (e.target.id === 'honorModal') {{
                closeHonorModal();
            }}
        }}

        function setHonorMode(mode) {{
            currentHonorMode = mode;
            const btnClass = document.getElementById('btnModeClass');
            const btnSection = document.getElementById('btnModeSection');
            const dropdownWrap = document.getElementById('sectionDropdownWrap');

            if (mode === 'class') {{
                btnClass.classList.add('active');
                btnSection.classList.remove('active');
                dropdownWrap.style.display = 'none';
            }} else {{
                btnSection.classList.add('active');
                btnClass.classList.remove('active');
                dropdownWrap.style.display = 'block';
            }}

            renderHonorToppers();
        }}

        function renderHonorToppers() {{
            const container = document.getElementById('podiumContainer');
            container.innerHTML = '';

            let list = [...DB.students];

            if (currentHonorMode === 'section') {{
                const selectedSec = document.getElementById('sectionSelect').value;
                if (selectedSec) {{
                    list = list.filter(st => st.class_name === selectedSec);
                }}
            }}

            const posKey = (currentHonorMode === 'class') ? 'class_pos' : 'sec_pos';

            const p1Students = list.filter(st => String(st[posKey]).trim() === '1');
            const p2Students = list.filter(st => String(st[posKey]).trim() === '2');
            const p3Students = list.filter(st => String(st[posKey]).trim() === '3');'''

new_modal_js = '''function openHonorModal() {{
            const secDropdown = document.getElementById('sectionSelect');
            secDropdown.innerHTML = '';
            DB.sections.forEach(sec => {{
                const opt = document.createElement('option');
                opt.value = sec;
                opt.textContent = sec;
                secDropdown.appendChild(opt);
            }});

            const clsDropdown = document.getElementById('classSelect');
            if (clsDropdown) {{
                clsDropdown.innerHTML = '';
                DB.grade_groups.forEach(grp => {{
                    const opt = document.createElement('option');
                    opt.value = grp;
                    opt.textContent = grp;
                    clsDropdown.appendChild(opt);
                }});
            }}

            if (currentStudent) {{
                secDropdown.value = currentStudent.class_name;
                if (clsDropdown) clsDropdown.value = currentStudent.grade_group;
            }}

            renderHonorToppers();
            document.getElementById('honorModal').style.display = 'flex';
        }}

        function closeHonorModal() {{
            document.getElementById('honorModal').style.display = 'none';
        }}

        function handleBackdropClick(e) {{
            if (e.target.id === 'honorModal') {{
                closeHonorModal();
            }}
        }}

        function setHonorMode(mode) {{
            currentHonorMode = mode;
            const btnClass = document.getElementById('btnModeClass');
            const btnSection = document.getElementById('btnModeSection');
            const secWrap = document.getElementById('sectionDropdownWrap');
            const clsWrap = document.getElementById('classDropdownWrap');

            if (mode === 'class') {{
                btnClass.classList.add('active');
                btnSection.classList.remove('active');
                if (secWrap) secWrap.style.display = 'none';
                if (clsWrap) clsWrap.style.display = 'block';
            }} else {{
                btnSection.classList.add('active');
                btnClass.classList.remove('active');
                if (secWrap) secWrap.style.display = 'block';
                if (clsWrap) clsWrap.style.display = 'none';
            }}

            renderHonorToppers();
        }}

        function renderHonorToppers() {{
            const container = document.getElementById('podiumContainer');
            container.innerHTML = '';

            let list = [...DB.students];

            if (currentHonorMode === 'section') {{
                const selectedSec = document.getElementById('sectionSelect').value;
                if (selectedSec) {{
                    list = list.filter(st => st.class_name === selectedSec);
                }}
            }} else {{
                const clsDropdown = document.getElementById('classSelect');
                const selectedClass = clsDropdown ? clsDropdown.value : '';
                if (selectedClass) {{
                    list = list.filter(st => st.grade_group === selectedClass);
                }}
            }}

            const posKey = (currentHonorMode === 'class') ? 'class_pos' : 'sec_pos';

            let p1Students = list.filter(st => String(st[posKey]).trim() === '1');
            let p2Students = list.filter(st => String(st[posKey]).trim() === '2');
            let p3Students = list.filter(st => String(st[posKey]).trim() === '3');

            // Fallback if class_pos is not numbered 1, 2, 3
            if (p1Students.length === 0 && p2Students.length === 0 && p3Students.length === 0 && list.length > 0) {{
                const sorted = [...list].sort((a, b) => parseFloat(b.percentage || 0) - parseFloat(a.percentage || 0));
                p1Students = sorted.slice(0, 1);
                p2Students = sorted.slice(1, 2);
                p3Students = sorted.slice(2, 3);
            }}'''

if old_modal_js in text:
    text = text.replace(old_modal_js, new_modal_js)
    print("Updated modal JS logic.")
else:
    print("Warning: old_modal_js not matched directly!")

with open(r"C:\Users\IT_OFFICE02\Desktop\Chenab_Result_Portal\build_desktop_portal.py", "w", encoding="utf-8") as f:
    f.write(text)

print("Saved polish updates to build_desktop_portal.py.")
