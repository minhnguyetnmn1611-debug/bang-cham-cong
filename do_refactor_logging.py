import os
import re

files_to_check = ['app.py', 'page_mos.py', 'page_checkin.py', 'page_gamification.py', 'db.py', 'excel_export.py', 'utils.py']

# For each file, we will add `from log_config import logger` at the top (after other imports)
# And replace `print(f"Lỗi: {e}")` and similar variants with `logger.warning(f"Lỗi: {e}", exc_info=True)`

def process_file(filepath):
    if not os.path.exists(filepath): return
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()
    
    original_content = content
    
    # 1. Add import if not present
    if "from log_config import logger" not in content:
        # Find a good place to insert (after standard library imports or streamlit)
        if "import streamlit as st" in content:
            content = content.replace("import streamlit as st", "import streamlit as st\nfrom log_config import logger", 1)
        elif "import pandas as pd" in content:
            content = content.replace("import pandas as pd", "import pandas as pd\nfrom log_config import logger", 1)
        elif "import sqlite3" in content:
            content = content.replace("import sqlite3", "import sqlite3\nfrom log_config import logger", 1)
        else:
            content = "from log_config import logger\n" + content

    # 2. Replace print(f"Lỗi: {e}") with logger
    # There are multiple variants, let's use regex
    # Match: except Exception as e: print(f"Lỗi: {e}");
    # We will replace just the print part.
    
    def replacer(match):
        exc_var = match.group(1) # usually e
        # If it's in db or excel_export, maybe use logger.error?
        # Since it's 50+ places, warning is safer for flow control.
        # But we can just use logger.warning with exc_info=True
        # Wait, if it's `except Exception: print(...)` without `as e`, that's rare but possible.
        return f'except Exception as {exc_var}: logger.warning(f"Lỗi: {{{exc_var}}}", exc_info=True);'

    # Match `except Exception as e: print(f"Lỗi: {e}");`
    content = re.sub(r'except Exception as (\w+):\s*print\([fF]?"[L|l]ỗi:?\s*\{?\1\}?"\);?', replacer, content)
    
    # Also match `except Exception as e:\n    print(f"Lỗi: {e}")`
    def replacer2(match):
        exc_var = match.group(1)
        spaces = match.group(2)
        return f'except Exception as {exc_var}:\n{spaces}logger.warning(f"Lỗi: {{{exc_var}}}", exc_info=True)'
    
    content = re.sub(r'except Exception as (\w+):\n(\s+)print\([fF]?"[L|l]ỗi:?\s*\{?\1\}?"\)', replacer2, content)
    
    # Also match `except Exception:\n` where they didn't catch `e`
    # We'll just ignore those for now or manually fix them if there are few.
    
    # Let's also find print statements inside except blocks broadly
    def replacer3(match):
        spaces = match.group(1)
        var = match.group(2)
        return f'{spaces}logger.warning(f"Lỗi: {{{var}}}", exc_info=True)'
    content = re.sub(r'(\s+)print\([fF]?"[L|l]ỗi:?\s*\{(\w+)\}?"\)', replacer3, content)

    if content != original_content:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"Updated {filepath}")

for f in files_to_check:
    process_file(f)
