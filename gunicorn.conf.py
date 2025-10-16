"""
Gunicorn configuration for the AMI calculator deployment.

Render launches the service with `gunicorn app:app`. Keeping this
file in the repository lets us bump worker timeouts without
touching any optimization logic.
"""

timeout = 180
graceful_timeout = 180
