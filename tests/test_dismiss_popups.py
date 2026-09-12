from tools.dismiss_popups import dismiss_script


def test_dismiss_script_covers_cookie_buttons():
    js = dismiss_script()
    for needle in ("принять все", "allow all", "allow essential", "accept all", "отклонить"):
        assert needle in js.lower() or needle in js
