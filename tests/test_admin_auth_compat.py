import app as app_module


def test_admin_endpoints_accept_api_key_in_compat_mode():
    client = app_module.app.test_client()

    old_api = app_module.API_KEY
    old_admin = app_module.ADMIN_KEY
    old_allow = app_module.ALLOW_API_KEY_FOR_ADMIN

    try:
        app_module.API_KEY = "api-key-123"
        app_module.ADMIN_KEY = "admin-key-456"
        app_module.ALLOW_API_KEY_FOR_ADMIN = True

        resp = client.get("/api/rent-calculators", headers={"X-API-Key": "api-key-123"})
        assert resp.status_code == 200
    finally:
        app_module.API_KEY = old_api
        app_module.ADMIN_KEY = old_admin
        app_module.ALLOW_API_KEY_FOR_ADMIN = old_allow


def test_admin_endpoints_reject_api_key_when_compat_disabled():
    client = app_module.app.test_client()

    old_api = app_module.API_KEY
    old_admin = app_module.ADMIN_KEY
    old_allow = app_module.ALLOW_API_KEY_FOR_ADMIN

    try:
        app_module.API_KEY = "api-key-123"
        app_module.ADMIN_KEY = "admin-key-456"
        app_module.ALLOW_API_KEY_FOR_ADMIN = False

        resp = client.get("/api/rent-calculators", headers={"X-API-Key": "api-key-123"})
        assert resp.status_code == 401
    finally:
        app_module.API_KEY = old_api
        app_module.ADMIN_KEY = old_admin
        app_module.ALLOW_API_KEY_FOR_ADMIN = old_allow
