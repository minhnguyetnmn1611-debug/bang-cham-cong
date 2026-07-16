import re

file_path = 'app.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

lines = content.split('\n')
for i in range(2730, 3430):
    if 'st.markdown("""' in lines[i]:
        lines[i] = lines[i].replace('st.markdown("""', 'st.markdown(f"""')

with open(file_path, 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))
print('Fixed f-strings!')
