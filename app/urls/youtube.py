from django.urls import path
from ..views import (
    YouTubeTranscriptView,

)

urlpatterns = [
    path('tools/youtube-transcript', YouTubeTranscriptView.as_view(), name='youtube-transcript'),
]
