import io

EDITS = {
    "source/test_delta_lag.py": [
        (
            'self.patch.object(self.delta_main, "_bulk_owning_manual", True)',
            'self.patch.object(self.delta_state, "_bulk_owning_manual", True)',
        ),
        (
            "        import pandas as pd\n        import main as delta_main\n",
            "        import delta_state\n        import pandas as pd\n        import main as delta_main\n",
        ),
        (
            "        self.delta_main = delta_main\n        self.pd = pd\n",
            "        self.delta_main = delta_main\n        self.delta_state = delta_state\n        self.pd = pd\n",
        ),
    ],
    "source/test_calc_mode_coexistence.py": [
        ("delta_main._manual_since", "delta_state._manual_since"),
        ("delta_main._last_observed_calc_raw", "delta_state._last_observed_calc_raw"),
        ("delta_main._calc_churn_count", "delta_state._calc_churn_count"),
        ("import main as delta_main", "import delta_state\nimport main as delta_main"),
    ],
}

for path, pairs in EDITS.items():
    text = io.open(path, encoding="utf-8").read()
    for old, new in pairs:
        if old not in text:
            print("MISSING in", path, ":", old[:70])
        text = text.replace(old, new)
    io.open(path, "w", encoding="utf-8", newline="\n").write(text)
    print("patched", path)
