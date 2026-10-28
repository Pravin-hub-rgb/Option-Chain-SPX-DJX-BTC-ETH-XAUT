import io

EDITS = {
    "source/tests/test_00_isolation.py": [
        (
            "    def test_delta_log_dir_override_is_honoured(self):\n        import main\n        self.assertEqual(_LOG_ROOT, main.FEED_LOG_DIR)",
            "    def test_delta_log_dir_override_is_honoured(self):\n        import delta_diagnostics\n        self.assertEqual(_LOG_ROOT, delta_diagnostics.FEED_LOG_DIR)",
        ),
        (
            "    def test_delta_log_dir_is_not_the_customer_directory(self):\n        import main\n        self.assertNotEqual(\n            os.path.join(os.environ.get(\"LOCALAPPDATA\", \".\"), \"DeltaOptionChain\"),\n            main.FEED_LOG_DIR,\n        )",
            "    def test_delta_log_dir_is_not_the_customer_directory(self):\n        import delta_diagnostics\n        self.assertNotEqual(\n            os.path.join(os.environ.get(\"LOCALAPPDATA\", \".\"), \"DeltaOptionChain\"),\n            delta_diagnostics.FEED_LOG_DIR,\n        )",
        ),
    ],
    "source/tests/test_delta_startup.py": [
        ("src = inspect.getsource(main._lock_path)", "src = inspect.getsource(delta_instance._lock_path)"),
        (
            "src = inspect.getsource(main._take_over_locked_instance)",
            "src = inspect.getsource(delta_instance._take_over_locked_instance)",
        ),
        (
            "src = inspect.getsource(main.kill_previous_instances)",
            "src = inspect.getsource(delta_instance.kill_previous_instances)",
        ),
        (
            "import main\n\n\nclass ProductNamingTests",
            "import delta_instance\nimport main\n\n\nclass ProductNamingTests",
        ),
        (
            "    def test_product_exe_is_named_for_the_product(self):\n        self.assertEqual(main.PRODUCT_EXE, \"delta_option_chain.exe\")",
            "    def test_product_exe_is_named_for_the_product(self):\n        import delta_state\n\n        self.assertEqual(delta_state.PRODUCT_EXE, \"delta_option_chain.exe\")",
        ),
        (
            "        src = inspect.getsource(main.kill_previous_instances)\n        self.assertIn(\"PRODUCT_EXE\", src)",
            "        src = inspect.getsource(delta_instance.kill_previous_instances)\n        self.assertIn(\"PRODUCT_EXE\", src)",
        ),
        (
            "        src = inspect.getsource(main._lock_path)\n        self.assertIn(\"BASE_DIR\", src)",
            "        src = inspect.getsource(delta_instance._lock_path)\n        self.assertIn(\"BASE_DIR\", src)",
        ),
    ],
}

for path, pairs in EDITS.items():
    text = io.open(path, encoding="utf-8").read()
    for old, new in pairs:
        if old not in text:
            print("MISSING in", path, ":", old[:70].replace("\n", " | "))
        text = text.replace(old, new)
    io.open(path, "w", encoding="utf-8", newline="\n").write(text)
    print("patched", path)
