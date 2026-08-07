# V10.1 — API bağlantı göstergesi

Bu güncelleme API bağlantı durumlarını dosya yüklemeden önce, sayfanın üstünde görünür hale getirir.

## Değişenler

- Tavily / Abstract / Lusha için ayrı bağlantı kartları.
- Lusha secret'ı bulunamazsa açık hata mesajı.
- Anahtar değeri gösterilmeden secret adının tanınması.
- `LUSHA_API_KEY` yanında yanlışlıkla kullanılan `LUSHA_KEY` ve `lusha_api_key` adlarını da tanıyan geri uyumluluk.
- Sayfada ve sol panelde sürüm: `v10.1`.

## Kurulum

GitHub repository kökündeki `app.py` dosyasını bu paketteki `app.py` ile değiştirip commit et.
Streamlit Cloud'da **Reboot app** yap.

Secrets önerilen biçim:

```toml
TAVILY_API_KEY = "tvly-..."
ABSTRACT_API_KEY = "..."
LUSHA_API_KEY = "..."
APP_PASSWORD = "..."
```
