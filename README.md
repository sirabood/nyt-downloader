
# NYT Downloader

![NYT Downloader Logo](/Assets/logo.ico)

**NYT Downloader** is a powerful, user-friendly desktop utility that allows you to download videos, music, and playlists from **YouTube**, **YouTube Music**, **JioSaavn**, and many other [supported websites](https://github.com/yt-dlp/yt-dlp/blob/master/supportedsites.md). It leverages the advanced functionality of [yt-dlp](https://github.com/yt-dlp/yt-dlp) as its backend, with a sleek and intuitive **Tkinter-based GUI** frontend.

## 🚀 Features

- 🎞️ Download videos, audio, playlists, or entire channels
- 🌐 Supports YouTube, YouTube Music, JioSaavn, and [1000+ websites](https://github.com/yt-dlp/yt-dlp/blob/master/supportedsites.md)
- ⚙️ Full access to yt-dlp's advanced configuration settings like interconversion in different formats, download subtitle, metadata, comments, and much more
- 🎛️ Preset management with save/load functionality
- 🧾 Real-time log output and download queue management

## 🖼️ Screenshot
![Main App](/Screenshots/Main.png)
![Full settings 2](/Screenshots/Full%20settings%20(2).png)
![Full settings 3](/Screenshots/Full%20settings%20(3).png)
![Full settings 4](/Screenshots/Full%20settings%20(4).png)
![Full settings 5](/Screenshots/Full%20settings%20(5).png)
![Full settings 6](/Screenshots/Full%20settings%20(6).png)

## 🛠️ Requirements

- Python 3.8+
- [yt-dlp](https://github.com/yt-dlp/yt-dlp)
- [FFmpeg](https://www.ffmpeg.org/download.html) 
- `tkinter` (usually comes with Python)
- Recommended: Windows OS (others may work with modification)

## 📦 Installation

If you are a Windows user, you can download the pre-built [EXE-ZIP](https://github.com/Maurya-Nitin/nyt-downloader/releases/download/v1.0.0/NYT.Downloader.zip) package from the [Releases](https://github.com/Maurya-nitin/nyt-downloader/releases) page for a hassle-free setup.

If you are a Linux (debian), you can download the package [NYT Downloader_linux](https://github.com/Maurya-Nitin/nyt-downloader/releases/download/v1.0.0/NYT.Downloader_linux.zip).

Otherwise if you want to run it from source code, follow these steps(Linux/Mac users can give it a try):

1. **Clone the repository:**
   ```bash
   git clone https://github.com/Maurya-nitin/nyt-downloader.git
   cd nyt-downloader
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application:**
   ```bash
   python nyt_downloader.py
   ```

## 🔒 License

This project is released into the public domain under the terms of the [Unlicense](https://unlicense.org). You are free to use, modify, and distribute this software without any restrictions. For more details, see the [LICENSE](LICENSE) file.

## 👨‍💻 Authors

- **Nitin** (Lead Developer)
- Contributors welcome! Feel free to open issues or submit pull requests.

## 📫 Contact

For suggestions, bugs, or feedback, feel free to open an [Issue](https://github.com/Maurya-nitin/nyt-downloader/issues) on GitHub.

---

> NYT Downloader is not affiliated with YouTube, JioSaavn, or any other supported platforms. Please ensure you comply with their terms of service when using this tool.
