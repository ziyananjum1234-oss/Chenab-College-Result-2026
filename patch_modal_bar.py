with open(r"C:\Users\IT_OFFICE02\Desktop\Chenab_Result_Portal\build_desktop_portal.py", "r", encoding="utf-8") as f:
    text = f.read()

m1 = text.find('<div class="honor-control-bar">')
m2 = text.find('<!-- Podium List -->')

new_bar = '''<div class="honor-control-bar">
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
                </div>

                '''

text = text[:m1] + new_bar + text[m2:]

with open(r"C:\Users\IT_OFFICE02\Desktop\Chenab_Result_Portal\build_desktop_portal.py", "w", encoding="utf-8") as f:
    f.write(text)

print("Updated honor-control-bar cleanly.")
