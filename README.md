# Stego-App
Windows Steganography App


## About

**Stego App** allows you to embed hidden messages into 2D digital images securely.  

### Data Protection
The embedded data is protected using a combination of spiral traversal patterns (Corner, Direction, Depth), AES-based encryption, and compression.  
This makes the hidden data very hard to detect or extract without authorization.

- **Corner:** Defines the starting point of embedding: Top-Left (A), Top-Right (B), Bottom-Right (C), Bottom-Left (D).  
- **Depth:** Value `0` or `1` specifies whether embedding starts from the center or from the image edges.  
- **Compression:** Compression serves two purposes:
  * General purpose: reduce the data size before embedding.
  * Security enhancement: once applied, extraction without the correct decryption key will not produce any valid output.  
- **Encryption (confidential):** Final protection uses user-provided password and AES key size. The larger the key, the more random and secure the result becomes.  

### Analysis & Logging
The app records spiral coordinate traversal and common errors into a log file.  
If you encounter errors or bugs, please report them via **email: hersaputrayoda@gmail.com** or **[GitHub Issues](https://github.com/yoda24/Stego-App/issues)**.


---

## Supported Image Formats

✅ **Lossless (read & write):** PNG, BMP (24-bit), TIF, TIFF  
⚠️ **Lossy (JPEG/JPG):** Supported for reading & embedding, **but cannot be saved back** to lossy format because compression alters LSB data. Always save as PNG/BMP/TIFF after embedding.  

### Supported Bit Depths

| Bit Depth | Supported Modes             | Channels |
|-----------|-----------------------------|----------|
| 8-bit     | Grayscale                   | 1        |
| 16-bit    | Grayscale, RGB, RGBA        | 1/3/4    |
| 24-bit    | RGB                         | 3        |
| 32-bit    | RGBA                        | 4        |

---

## Capacity Estimation

The app estimates how much data can be embedded based on image dimensions and channels.  
Some metadata overhead is reserved: **flag (4 bit), length (32 bit), CRC (32 bit)**.

Example:

| Width | Height | Channels | Raw Capacity (bits) | Final Capacity (after overhead) |
|-------|--------|----------|---------------------|---------------------------------|
| 10    | 10     | 3        | 300                 | 232 bits (29 bytes)             |
| 5000  | 5000   | 4        | 100,000,000         | 99,999,932 bits (~12.50 MB)     |
| 16000 | 16000  | 4        | 1,024,000,000       | 1,023,999,932 bits (~122 MB)    |

**Maximum capacity:**  
The length header uses **4 bytes (32-bit)**, meaning the maximum storable size is:  
`0xFFFFFFFF = 4,294,967,295 bytes` (~4 GB).  

---

## System Requirements

- **Tested on:** Windows 10 (64-bit)  
- **Minimum RAM**: 4 GB (8 GB recommended for large images)
---

## Download & Run

1. Go to [Releases](https://github.com/yoda24/Stego-App/releases)  
2. Download the latest `.exe` build  
3. Run the application directly — no installation required  

---

## Built With

- Python 3.1 (64-bit)  
- imageio  
- numpy  
- PyQt6  
- PyInstaller (for packaging)  

---

## License

This project is licensed under the MIT License – see the [LICENSE](LICENSE) file for details.

---

## Support

☕ If you like this project, you can support me:  
- [PayPal](https://paypal.me/yodahs)  
- [Saweria](https://saweria.co/digishop)  

---
