# Store Review Checker — Streamlit V2

Google Play ve Apple App Store linklerini toplu kontrol eder.

## Çalıştırma

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

App Store yorumları Apple'ın müşteri yorumları Atom XML akışından okunur. Akış erişilemez veya hız sınırına takılırsa uygulama bunu "yorum yok" olarak değil, "bilinmiyor/kısmi" olarak gösterir.
