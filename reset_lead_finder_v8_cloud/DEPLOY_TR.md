# Reset Lead Finder Cloud kurulumu

## Önerilen yapı

Streamlit uygulaması sürekli çalışan Python süreci ve WebSocket bağlantısı kullanır. Bu nedenle uygulamanın kendisi Streamlit Community Cloud veya Render üzerinde çalışır. Netlify klasörü yalnızca markalı giriş/yönlendirme sayfasıdır.

## A) Streamlit Community Cloud — en kolay yöntem

1. GitHub'da PRIVATE bir repository oluştur: `reset-lead-finder`.
2. Bu klasördeki dosyaları repository köküne yükle. `netlify_gateway` klasörü de kalabilir.
3. `share.streamlit.io` üzerinden GitHub hesabınla giriş yap.
4. **Create app** seç ve repository/branch/main file olarak `app.py` belirle.
5. App settings > Secrets alanına aşağıdakini kendi anahtarlarınla ekle:

```toml
TAVILY_API_KEY = "tvly-..."
ABSTRACT_API_KEY = "..."
APP_PASSWORD = "uzun-ve-guclu-bir-sifre"
```

6. Deploy et. Oluşan URL örneği: `https://reset-lead-finder.streamlit.app`.
7. Anahtarları hiçbir zaman GitHub dosyalarına yazma.

## B) Netlify giriş adresi

1. `netlify_gateway/config.js` dosyasını aç.
2. `https://YOUR-APP.streamlit.app` alanını gerçek Streamlit URL'siyle değiştir.
3. Yalnızca `netlify_gateway` klasörünü Netlify Drop alanına sürükle.
4. Netlify adresini istersen `lead.resetiletisim.com` gibi bir subdomain'e bağla.

Netlify sayfası uygulamayı barındırmaz; güvenli ve markalı giriş bağlantısı sağlar.

## C) Render alternatifi

Repository'yi Render'a bağla. `render.yaml` otomatik yapılandırmayı içerir. Environment bölümüne `TAVILY_API_KEY`, `ABSTRACT_API_KEY` ve `APP_PASSWORD` ekle.

## Yerel test

Windows'ta `start_windows.bat` dosyasını çalıştır. Bulut secrets yoksa uygulama API anahtarlarını oturum içinde girmeni ister.

## Güvenlik

- Repository private kalsın.
- Gerçek API anahtarını `.streamlit/secrets.toml.example` içine yazma.
- Uygulama parolasını güçlü ve farklı seç.
- Yüklenen Excel dosyaları uygulama tarafından kalıcı diske kaydedilmez; sonuç kullanıcı tarafından indirilir.
