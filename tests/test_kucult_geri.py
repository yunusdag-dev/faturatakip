# -*- coding: utf-8 -*-
"""Uygulamayı küçültüp görev çubuğundan geri yüklemek (Win+D / Alt+Tab) modal pencereyle birlikte çalışmalı.

Gerçek uygulamayı ayrı bir süreçte açar (yaklaşık 30 sn); yalnız Windows ve ekran gerekir."""
import os
import subprocess
import sys
import unittest


def _tk_var():
    try:
        import tkinter
        tkinter.Tk().destroy()
        return True
    except Exception:
        return False


@unittest.skipUnless(os.name == "nt" and _tk_var(), "yalnız Windows + ekran")
class KucultGeriYukle(unittest.TestCase):
    def test_modal_ve_diger_pencereler_kucult_geri_yukle(self):
        yardimci = os.path.join(os.path.dirname(os.path.abspath(__file__)), "yardimci_kucult_geri.py")
        p = subprocess.run([sys.executable, "-I", yardimci], capture_output=True, text=True, encoding="utf-8",
                           errors="replace", timeout=200)
        self.assertEqual(p.returncode, 0, "\n" + p.stdout[-3500:] + p.stderr[-1500:])


if __name__ == "__main__":
    unittest.main()
