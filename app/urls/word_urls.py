from django.urls import path
from ..views import (
    WordToJPGView,
    WordToTXTView,
    TXTToWordView,
    MergeDOCXView,
    SplitDOCXView,
    WordToMarkdownView,
    MarkdownToWordView,
)

urlpatterns = [
    path("tools/word-to-jpg", WordToJPGView.as_view(), name="word-to-jpg"),
    path("tools/word-to-txt", WordToTXTView.as_view(), name="word-to-txt"),
    path("tools/txt-to-word", TXTToWordView.as_view(), name="txt-to-word"),
    path("tools/merge-docx", MergeDOCXView.as_view(), name="merge-docx"),
    path("tools/split-docx", SplitDOCXView.as_view(), name="split-docx"),
    path("tools/word-to-markdown", WordToMarkdownView.as_view(), name="word-to-markdown"),
    path("tools/markdown-to-word", MarkdownToWordView.as_view(), name="markdown-to-word"),
]
