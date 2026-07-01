import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parent

def run(cmd):
    print('> ' + ' '.join(cmd))
    subprocess.check_call(cmd, cwd=ROOT)

run(['python', 'build_bundle.py'])
run(['python', 'scripts/grab_go_patch.py'])
run(['python', 'scripts/grab_go_cloud_patch.py'])
run(['python', 'scripts/grab_go_search_patch.py'])
run(['python', 'scripts/main_site_cloud_patch.py'])

if any((ROOT / 'input' / 'anagrafica').glob('*.xlsx')):
    for script in [
        'scripts/planning_standalone.py',
        'scripts/planning_nadir_alias_patch.py',
        'scripts/planning_excel_stable_patch.py',
        'scripts/planning_ui_mobile_patch.py',
        'scripts/planning_ui_cleanup_patch.py',
        'scripts/planning_edit_ui_patch.py',
        'scripts/planning_edit_scroll_fix_patch.py',
        'scripts/planning_edit_scroll_hardfix_patch.py',
        'scripts/planning_edit_sticky_save_patch.py',
        'scripts/planning_edit_dayblock_patch.py',
        'scripts/planning_start_home_patch.py',
        'scripts/planning_manage_nadir_alias_patch.py',
        'scripts/planning_grab_visit_mode_patch.py',
        'scripts/planning_period_json_patch.py',
    ]:
        run(['python', script])
else:
    print('Planning standalone saltato: nessun file .xlsx in input/anagrafica')

print('Build Cloudflare Pages completata. Cartella output: docs')
