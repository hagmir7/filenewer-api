from rest_framework import serializers


class PDFUploadSerializer(serializers.Serializer):
    """Validates the incoming multipart upload."""

    file = serializers.FileField(help_text="PDF file to convert (max 50 MB).")

    def validate_file(self, value):
        # Check MIME type / extension
        name = value.name.lower()
        content_type = getattr(value, "content_type", "")

        if not name.endswith(".pdf") and "pdf" not in content_type:
            raise serializers.ValidationError(
                "Only PDF files are accepted. "
                f"Received: '{value.name}' (content-type: {content_type})"
            )

        # Size guard (50 MB)
        max_bytes = 50 * 1024 * 1024
        if value.size > max_bytes:
            mb = value.size / (1024 * 1024)
            raise serializers.ValidationError(
                f"File too large ({mb:.1f} MB). Maximum allowed size is 50 MB."
            )

        return value
