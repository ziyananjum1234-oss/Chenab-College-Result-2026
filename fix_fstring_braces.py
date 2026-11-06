with open(r"C:\Users\IT_OFFICE02\Desktop\Chenab_Result_Portal\build_desktop_portal.py", "r", encoding="utf-8") as f:
    text = f.read()

bad_block = '''if (document.getElementById('dispClass')) {
                document.getElementById('dispClass').textContent = st.class_name;
            }

            const secHeadImg = document.getElementById('secHeadSigImg');
            const targetSig = st.sec_head_sig ? SIGNATURES[st.sec_head_sig] : (SIGNATURES.sec_sajida_4_8_g || '');
            if (secHeadImg && targetSig) {
                secHeadImg.src = targetSig;
            }
            document.getElementById('secHeadName').textContent = st.sec_head_name || 'Section Head';
            document.getElementById('secHeadTitle').textContent = st.sec_head_title || 'Chenab College Shorkot';'''

good_block = '''if (document.getElementById('dispClass')) {{
                document.getElementById('dispClass').textContent = st.class_name;
            }}

            const secHeadImg = document.getElementById('secHeadSigImg');
            const targetSig = st.sec_head_sig ? SIGNATURES[st.sec_head_sig] : (SIGNATURES.sec_sajida_4_8_g || '');
            if (secHeadImg && targetSig) {{
                secHeadImg.src = targetSig;
            }}
            document.getElementById('secHeadName').textContent = st.sec_head_name || 'Section Head';
            document.getElementById('secHeadTitle').textContent = st.sec_head_title || 'Chenab College Shorkot';'''

if bad_block in text:
    text = text.replace(bad_block, good_block)
    print("Fixed f-string braces.")
else:
    print("bad_block not matched!")

with open(r"C:\Users\IT_OFFICE02\Desktop\Chenab_Result_Portal\build_desktop_portal.py", "w", encoding="utf-8") as f:
    f.write(text)
