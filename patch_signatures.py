import os
import re

build_script_path = r"C:\Users\IT_OFFICE02\Desktop\Chenab_Result_Portal\build_desktop_portal.py"
with open(build_script_path, "r", encoding="utf-8") as f:
    code = f.read()

# 1. Add signature loading after font loading
sig_loader = '''
# 3. Embed Digital Signatures
sig_dir = os.path.join(portal_folder, "signatures")
sig_map = {
    "exam_head_sibtain": "exam_head_sibtain.png",
    "sec_shoukat_9_12_b": "sec_shoukat_9_12_b.png",
    "sec_sambeen_9_12_g": "sec_sambeen_9_12_g.png",
    "sec_sajida_4_8_g": "sec_sajida_4_8_g.png",
    "sec_abbas_4_8_b": "sec_abbas_4_8_b.png",
    "sec_farhat_1_3": "sec_farhat_1_3.png"
}
signatures_b64 = {}
for k, fname in sig_map.items():
    p = os.path.join(sig_dir, fname)
    if os.path.exists(p):
        with open(p, "rb") as sf:
            signatures_b64[k] = "data:image/png;base64," + base64.b64encode(sf.read()).decode("utf-8")
    else:
        signatures_b64[k] = ""

signatures_json_str = json.dumps(signatures_b64)
'''

if "signatures_json_str = json.dumps(signatures_b64)" not in code:
    anchor = 'font_b64 = base64.b64encode(f.read()).decode("utf-8")'
    if anchor in code:
        code = code.replace(anchor, anchor + "\n" + sig_loader)
        print("Inserted signature loader.")
    else:
        print("Anchor not found for font loader!")

# 2. Add signature CSS
sig_css = '''
        .sig-col {{
            flex: 1;
            max-width: 240px;
            text-align: center;
        }}
        .sig-img-container {{
            height: 48px;
            display: flex;
            align-items: flex-end;
            justify-content: center;
            margin-bottom: 2px;
        }}
        .sig-img {{
            max-height: 45px;
            max-width: 140px;
            object-fit: contain;
            filter: contrast(1.08);
        }}
        .sig-rule {{
            width: 85%;
            margin: 0 auto 5px;
            border-top: 1px solid #64748b;
        }}'''

old_sig_col = '''.sig-col {{
            flex: 1;
            max-width: 240px;
            text-align: center;
        }}
        .sig-rule {{
            width: 85%;
            margin: 0 auto 5px;
            border-top: 1px solid #64748b;
        }}'''

if old_sig_col in code:
    code = code.replace(old_sig_col, sig_css)
    print("Replaced sig-col CSS.")
else:
    print("Old sig-col CSS not matched exactly, checking variants...")

# 3. Add print CSS for signatures
old_print_footer = '''.cert-footer-grid {{
                margin-top: 10px !important;
                padding-top: 8px !important;
            }}
            .sig-name {{'''

new_print_footer = '''.cert-footer-grid {{
                margin-top: 6px !important;
                padding-top: 5px !important;
            }}
            .sig-img-container {{
                height: 38px !important;
                margin-bottom: 1px !important;
            }}
            .sig-img {{
                max-height: 36px !important;
                max-width: 110px !important;
            }}
            .sig-name {{'''

if old_print_footer in code:
    code = code.replace(old_print_footer, new_print_footer)
    print("Updated print footer CSS.")

# 4. Update HTML footer with signature image elements
old_footer_html = '''                    <!-- Signatures: EXACTLY "Exam Head" & "Section Head" -->
                    <div class="cert-footer-grid">
                        <div class="sig-col">
                            <div class="sig-rule"></div>
                            <div class="sig-name">Exam Head</div>
                            <div class="sig-sub">Incharge Examination</div>
                        </div>
                        
                        <div class="official-seal">
                            <span class="seal-txt-top">Chenab College</span>
                            <span class="seal-txt-mid">VERIFIED</span>
                            <span class="seal-txt-bot">Exam Branch</span>
                        </div>

                        <div class="sig-col">
                            <div class="sig-rule"></div>
                            <div class="sig-name">Section Head</div>
                            <div class="sig-sub">Chenab College Shorkot</div>
                        </div>
                    </div>'''

new_footer_html = '''                    <!-- Signatures: EXACTLY "Exam Head" & Dynamic "Section Head" -->
                    <div class="cert-footer-grid">
                        <div class="sig-col">
                            <div class="sig-img-container">
                                <img id="examHeadSigImg" src="" class="sig-img" alt="Exam Head Signature">
                            </div>
                            <div class="sig-rule"></div>
                            <div class="sig-name" id="examHeadName">Sir Sibtain Haider</div>
                            <div class="sig-sub" id="examHeadTitle">Exam Head (Incharge Examination)</div>
                        </div>
                        
                        <div class="official-seal">
                            <span class="seal-txt-top">Chenab College</span>
                            <span class="seal-txt-mid">VERIFIED</span>
                            <span class="seal-txt-bot">Exam Branch</span>
                        </div>

                        <div class="sig-col">
                            <div class="sig-img-container">
                                <img id="secHeadSigImg" src="" class="sig-img" alt="Section Head Signature">
                            </div>
                            <div class="sig-rule"></div>
                            <div class="sig-name" id="secHeadName">Mrs. Sajida</div>
                            <div class="sig-sub" id="secHeadTitle">Section Head (Girls Wing 4th - 8th)</div>
                        </div>
                    </div>'''

if old_footer_html in code:
    code = code.replace(old_footer_html, new_footer_html)
    print("Updated HTML footer markup.")
else:
    print("Footer HTML not matched exactly!")

# 5. In JS: Add SIGNATURES constant and getSectionHeadInfo function
js_insert = '''
        const SIGNATURES = {signatures_json_str};

        function getSectionHeadInfo(className) {{
            const cls = (className || "").toUpperCase();
            
            // Check Class Level
            let is9to12 = false;
            let is4to8 = false;
            let is1to3 = false;
            
            if (/1ST\\s*YEAR|2ND\\s*YEAR|11TH|12TH|9TH|10TH|\\b9\\b|\\b10\\b|\\b11\\b|\\b12\\b|MATRIC|FSC|ICS|FA\\b/.test(cls)) {{
                is9to12 = true;
            }} else if (/4TH|5TH|6TH|7TH|8TH|\\b4\\b|\\b5\\b|\\b6\\b|\\b7\\b|\\b8\\b/.test(cls)) {{
                is4to8 = true;
            }} else if (/1ST|2ND|3RD|PREP|NURSERY|KG|\\b1\\b|\\b2\\b|\\b3\\b/.test(cls)) {{
                is1to3 = true;
            }} else {{
                is4to8 = true;
            }}

            // Gender Category: Girls (G) vs Boys (B)
            const isGirls = cls.includes('(G)') || cls.includes('GIRLS') || cls.includes('GIRL') || cls.includes('JADES');

            if (is1to3) {{
                return {{
                    name: "Ms. Farhat",
                    title: "Section Head (Classes 1st - 3rd)",
                    sig: SIGNATURES.sec_farhat_1_3
                }};
            }} else if (is4to8) {{
                if (isGirls) {{
                    return {{
                        name: "Mrs. Sajida",
                        title: "Section Head (Girls Wing 4th - 8th)",
                        sig: SIGNATURES.sec_sajida_4_8_g
                    }};
                }} else {{
                    return {{
                        name: "Mr. Abbas",
                        title: "Section Head (Boys Wing 4th - 8th)",
                        sig: SIGNATURES.sec_abbas_4_8_b
                    }};
                }}
            }} else if (is9to12) {{
                if (isGirls) {{
                    return {{
                        name: "Mrs. Sambeen Ivans",
                        title: "Section Head (Girls Wing 9th - 12th)",
                        sig: SIGNATURES.sec_sambeen_9_12_g
                    }};
                }} else {{
                    return {{
                        name: "Sir Shoukat Sb",
                        title: "Section Head (Boys Wing 9th - 12th)",
                        sig: SIGNATURES.sec_shoukat_9_12_b
                    }};
                }}
            }}

            return {{
                name: isGirls ? "Mrs. Sajida" : "Mr. Abbas",
                title: isGirls ? "Section Head (Girls Wing)" : "Section Head (Boys Wing)",
                sig: isGirls ? SIGNATURES.sec_sajida_4_8_g : SIGNATURES.sec_abbas_4_8_b
            }};
        }}
'''

js_anchor = "const DB = {json_str};"
if js_anchor in code:
    code = code.replace(js_anchor, js_anchor + "\n" + js_insert)
    print("Inserted SIGNATURES and getSectionHeadInfo into JS.")
else:
    print("js_anchor not found!")

# 6. Call in renderResult(st)
render_anchor = "document.getElementById('dispRemarks').textContent = st.remarks || 'Keep up the good work.';"
sig_render_call = '''document.getElementById('dispRemarks').textContent = st.remarks || 'Keep up the good work.';

            // Render Signatures
            const examHeadImg = document.getElementById('examHeadSigImg');
            if (examHeadImg && SIGNATURES.exam_head_sibtain) {{
                examHeadImg.src = SIGNATURES.exam_head_sibtain;
            }}
            document.getElementById('examHeadName').textContent = 'Sir Sibtain Haider';
            document.getElementById('examHeadTitle').textContent = 'Exam Head (Incharge Examination)';

            const secInfo = getSectionHeadInfo(st.class_name);
            const secHeadImg = document.getElementById('secHeadSigImg');
            if (secHeadImg && secInfo.sig) {{
                secHeadImg.src = secInfo.sig;
            }}
            document.getElementById('secHeadName').textContent = secInfo.name;
            document.getElementById('secHeadTitle').textContent = secInfo.title;'''

if render_anchor in code:
    code = code.replace(render_anchor, sig_render_call)
    print("Inserted signature rendering into renderResult.")
else:
    print("render_anchor not found!")

with open(build_script_path, "w", encoding="utf-8") as f:
    f.write(code)

print("Saved updated build_desktop_portal.py successfully!")
