from .pdf_urls import urlpatterns as pdf_urlpatterns
from .csv_urls import urlpatterns as csv_urlpatterns
from .json_urls import urlpatterns as json_urlpatterns
from .excel_urls import urlpatterns as excel_urlpatterns
from .word_urls import urlpatterns as word_urlpatterns
from .base64_urls import urlpatterns as base64_urlpatterns
from .uuid_urls import urlpatterns as uuid_urlpatterns
from .crypto_urls import urlpatterns as crypto_urlpatterns
from .password_urls import urlpatterns as password_urlpatterns
from .timestamp_urls import urlpatterns as timestamp_urlpatterns
from .compare_urls import urlpatterns as compare_urlpatterns
from .ocr_urls import urlpatterns as ocr_urlpatterns
from .health_urls import urlpatterns as health_urlpatterns

urlpatterns = (
    pdf_urlpatterns
    + csv_urlpatterns
    + json_urlpatterns
    + excel_urlpatterns
    + word_urlpatterns
    + base64_urlpatterns
    + uuid_urlpatterns
    + crypto_urlpatterns
    + password_urlpatterns
    + timestamp_urlpatterns
    + compare_urlpatterns
    + ocr_urlpatterns
    + health_urlpatterns
)
