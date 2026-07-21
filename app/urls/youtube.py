from django.urls import path
from ..views import (
    YouTubeTranscriptView,
     YouTubeDownloadView,
    YouTubeVideoInfoView,

)

urlpatterns = [
    path('tools/youtube-transcript', YouTubeTranscriptView.as_view(), name='youtube-transcript'),
    path('tools/youtube-download', YouTubeDownloadView.as_view(),  name='youtube-download'),
    path('tools/youtube-info',     YouTubeVideoInfoView.as_view(), name='youtube-info'),
]
