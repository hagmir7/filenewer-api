from django.urls import path
from ..views import (
    PasswordGeneratorView,
    PasswordStrengthView,
    PassphraseGeneratorView,
    HashGeneratorView,
    HashCompareView,
)

urlpatterns = [
    path("tools/password-generate", PasswordGeneratorView.as_view(), name="password-generate"),
    path("tools/password-strength", PasswordStrengthView.as_view(), name="password-strength"),
    path("tools/passphrase-generate", PassphraseGeneratorView.as_view(), name="passphrase-generate"),
    path("tools/hash-generate", HashGeneratorView.as_view(), name="hash-generate"),
    path("tools/hash-compare", HashCompareView.as_view(), name="hash-compare"),
]
