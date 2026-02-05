"""Vercel Python entrypoint.

Expose the Flask app from the repository root so Vercel can serve
all routes defined in ``app.py`` (including ``/api/*``).
"""

from app import app

