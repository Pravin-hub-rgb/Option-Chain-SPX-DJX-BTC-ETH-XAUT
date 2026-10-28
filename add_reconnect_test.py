import os

base = r'C:\Users\Pravin\Downloads\Telegram Desktop\new2-deltax-option-chain\deltax-option-chain'

# --- Delta test ---
p = os.path.join(base, 'Delta Option Chain', 'source', 'tests', 'test_calc_mode_coexistence.py')
with open(p, 'r') as f:
    content = f.read()

new_test = '''
class WatchdogReconnectTests(unittest.TestCase):
    """ensure_calculation_not_stranded should reconnect when _app_cache is None."""

    def test_reconnect_in_watchdog_source(self):
        import inspect
        import main as delta_main
        src = inspect.getsource(delta_main.ensure_calculation_not_stranded)
        self.assertIn('xlwings', src)
        self.assertIn('xw.apps', src)


'''

idx = content.rfind('if __name__')
content = content[:idx] + new_test + content[idx:]
with open(p, 'w') as f:
    f.write(content)
print('Delta reconnect test added')

# --- US test ---
p2 = os.path.join(base, 'US Option Chain', 'source', 'tests', 'test_us_lag.py')
with open(p2, 'a') as f:
    f.write('''
class WatchdogReconnectTests(unittest.TestCase):
    """ensure_calculation_not_stranded should reconnect when _app_cache is None."""

    def test_reconnect_in_us_watchdog(self):
        import inspect
        import us_excel
        src = inspect.getsource(us_excel.ensure_calculation_not_stranded)
        self.assertIn('bulk_app', src)
''')
print('US reconnect test added')
