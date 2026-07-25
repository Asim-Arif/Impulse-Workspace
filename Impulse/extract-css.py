import re

with open(r'C:\Users\asimj\.nuget\packages\radzen.blazor\8.3.2\staticwebassets\css\material-base.css', 'r', encoding='utf-8') as f:
    content = f.read()

root_match = re.search(r':root\s*\{([^}]+)\}', content)
if root_match:
    root_vars = root_match.group(1)
    dialog_vars = re.findall(r'--rz-dialog-[^:]*:\s*[^;]+;', root_vars)
else:
    dialog_vars = []

# Find all selectors that contain .rz-dialog
dialog_css = re.findall(r'((?:[^\n{]*\.rz-dialog[^\n{]*)+\s*\{[^}]*\})', content)

out_css = ':root {\n'
for var in dialog_vars:
    out_css += '  ' + var + '\n'
out_css += '}\n\n'

# Adding some base styles for dialog buttons
base_button_vars = re.findall(r'(--rz-button-[^:]*:\s*[^;]+;|--rz-border-radius:\s*[^;]+;|--rz-text-[^:]*:\s*[^;]+;|--rz-white:\s*[^;]+;|--rz-black:\s*[^;]+;|--rz-primary-*:.*?\;|--rz-secondary-*:.*?\;|--rz-info-*:.*?\;)', root_vars)
out_css += ':root {\n'
for var in set(base_button_vars):
    out_css += '  ' + var + '\n'
out_css += '}\n\n'


for css in dialog_css:
    out_css += css + '\n\n'

# Add basic button styles so the Confirm button isn't completely unstyled
button_css = re.findall(r'(\.rz-button\s*\{[^}]*\})', content)
if button_css:
    out_css += button_css[0] + '\n\n'
button_primary_css = re.findall(r'(\.rz-button\.rz-primary\s*\{[^}]*\})', content)
if button_primary_css:
    out_css += button_primary_css[0] + '\n\n'
button_light_css = re.findall(r'(\.rz-button\.rz-light\s*\{[^}]*\})', content)
if button_light_css:
    out_css += button_light_css[0] + '\n\n'

with open(r'C:\Visual Studio .Net\Impulse Workspace\Impulse Solution\Impulse\wwwroot\css\radzen-dialog.css', 'w', encoding='utf-8') as f:
    f.write(out_css)
