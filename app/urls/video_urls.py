from django.urls import path
from ..views import (
    AddVoiceToVideoView,
)

urlpatterns = [
    path(
        "tools/add-voice-to-video", AddVoiceToVideoView.as_view(), name="add-voice-to-video",
    ),
]
