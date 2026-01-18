# Console-Based Mini Excel System

*A lightweight, terminal-based spreadsheet application built with C#*

## 🌐 Language Selection | Dil Seçimi

**English** | [Türkçe](#türkçe-dokümantasyon)

---

## English Documentation

### 📋 Overview

This is a feature-rich console application that simulates a spreadsheet environment using C# fundamentals. The system provides core spreadsheet functionality including cell operations, data manipulation, and file persistence - all through an intuitive command-line interface.

### ✨ Key Features

- **Grid-based Data Storage**: Supports up to 10 columns (A-J) and 15 rows (1-15)
- **Dynamic Cell Management**: Store both string and integer values with automatic type tracking
- **Flexible Operations**: 16 powerful operations for data manipulation
- **Smart Display**: Automatically truncates long values (>5 characters) with "_" indicator
- **Error Handling**: Comprehensive validation with descriptive error messages
- **Data Persistence**: Auto-saves spreadsheet content to `spreadsheet.txt`

### 🚀 Getting Started

#### Prerequisites
- .NET SDK (any version supporting C#)
- Terminal/Command Prompt

#### Installation

1. Clone the repository:
```bash
git clone https://github.com/AliYigitOzudogru/TUI-Based-Excel-Program.git
cd TUI-Based-Excel-Program
```

2. Compile the program:
```bash
csc Program.cs
```

3. Run the application:
```bash
Program.exe
```

Or use your preferred IDE (Visual Studio, Rider, VS Code).

### 📖 Operations Guide

#### Basic Cell Operations

**1. AssignValue** - Assign data to a cell
```
>> AssignValue(C4, integer, 45)
>> AssignValue(B9, string, HelloWorld)
```

**2. ClearCell** - Clear specific cell
```
>> ClearCell(E8)
```

**3. ClearAll** - Clear entire spreadsheet
```
>> ClearAll()
```

**Cell Query** - View full content
```
>> E8
```

#### Structure Operations

**4. AddRow** - Insert new row
```
>> AddRow(5, up)
>> AddRow(8, down)
```

**5. AddColumn** - Insert new column
```
>> AddColumn(C, right)
>> AddColumn(D, left)
```

#### Copy Operations

**6. Copy** - Copy cell to another cell
```
>> Copy(E8, A8)
```

**7. CopyColumn** - Copy entire column
```
>> CopyColumn(E, A)
```

**8. CopyRow** - Copy entire row
```
>> CopyRow(3, 7)
```

#### Cut Operations

**9. X** - Cut and paste cell
```
>> X(E8, A8)
```

**10. XColumn** - Cut and paste column
```
>> XColumn(E, A)
```

**11. XRow** - Cut and paste row
```
>> XRow(3, 7)
```

#### Mathematical & String Operations

**12. Multiplication/Repetition (*)** 
- Integer × Integer = Multiplication
- String × Integer = Repetition (positive: normal, negative: reversed)
```
>> *(C4, A13, I14)        // 4 * 4 = 16
>> *(H14, H8, J1)         // "ABc" * 2 = "ABcABc"
>> *(I10, H14, A14)       // "ABc" * -2 = "cBAcBA"
```

**13. Addition/Concatenation (+)**
- Integer + Integer + Integer = Addition
- Any String involved = Concatenation (with case conversion)
```
>> +(A5, C4, D4)          // 6875 + 45 = 6920
>> +(E11, F11, G11, J6)   // Concatenation with case selection
```

**14. Division/String Slicing (/)**
- Integer ÷ Integer = Division
- String ÷ Integer = Take first/last portion
```
>> /(B9, H8, J1)          // "Algorithm" / 2 = "Algorith"
>> /(I10, B9, A14)        // "Algorithm" / -2 = "esruoc m"
```

**15. Subtraction/Character Removal (-)**
- Integer - Integer = Subtraction
- String - Integer = Remove ASCII character
- String - String = Remove substring occurrences
```
>> -(C4, I10, I14)        // 45 - (-2) = 47
>> -(G7, G6, J1)          // "CAATBA" - 65('A') = "CTB"
>> -(F3, D1, A14)         // "A3B4 CVB4E" - "B4" = "A3 CVE"
```

**16. Encryption (#)**
- Shifts ASCII values of string characters by integer amount
- Valid shift range: [-20, 30]
```
>> #(A1, A13, D5)         // "Apple" + 4 = "Ettpi"
>> #(H8, B9, J10)         // "Algorithm" + 2 = "Epksvmxlq"
```

### 🔧 Technical Implementation

#### Design Highlights

- **No External Parsing Libraries**: Custom input validation logic throughout
- **2D Array Architecture**: Core data structure for spreadsheet grid
- **Manual Type Checking**: Explicit validation for all operations without TryParse
- **Procedural Approach**: Organized through functions and procedures
- **Character-by-character Processing**: String manipulations implemented manually

#### Data Structure
```
- Main Grid: string[15,10]
- Type Tracker: string[15,10] 
- Initial Size: 8 rows × 5 columns
```

#### Error Handling
- Boundary validation for all operations
- Type compatibility checks
- Parameter count verification
- Custom error messages for each scenario

### 📝 Example Session

```
   A      B      C      D      E  
1|      |      |      |      |      |
2|      |      |      |      |      |
3|      |      |      |      |      |
4|      |      |      |      |      |
5|      |      |      |      |      |
6|      |      |      |      |      |
7|      |      |      |      |      |
8|      |      |      |      |      |

>> AssignValue(B2, string, HelloWorld)
Operation is done!

   A      B      C      D      E  
1|      |      |      |      |      |
2|      |Hello_|      |      |      |

>> B2
HelloWorld

>> AssignValue(C2, integer, 5)
>> *(B2, C2, D2)
Operation is done!

>> D2
HelloWorldHelloWorldHelloWorldHelloWorldHelloWorld
```

### 📄 File Output

On exit, the program saves to `spreadsheet.txt`:
```
Current spreadsheet size: 8x5
Cell data and types preserved for future sessions
```

### 🛡️ Error Messages

- `Operation is done!` - Success
- `Illegal position assignment!` - Invalid cell/row/column reference
- `Out of bounds exception!` - Exceeds maximum grid size
- `Illegal operation!` - Type mismatch or invalid parameter combination
- Custom messages for specific operation constraints

### 🎯 Constraints

- Maximum grid: 10 columns × 15 rows
- ASCII character range for (-) operation: [33, 126]
- Encryption shift range: [-20, 30]
- All mathematical operations require assigned cells (no unassigned operands)

### 👨‍💻 Development Notes

This project emphasizes fundamental programming concepts:
- Manual input parsing and validation
- Array manipulation and boundary management
- String processing without built-in helpers
- Procedural decomposition
- Error-first design philosophy

### 📜 License

This project is available for educational and personal use.

---

## Türkçe Dokümantasyon

### 📋 Genel Bakış

C# temel kavramlarıyla oluşturulmuş, zengin özelliklere sahip bir konsol uygulamasıdır. Sistem, sezgisel bir komut satırı arayüzü üzerinden hücre işlemleri, veri manipülasyonu ve dosya kalıcılığı dahil olmak üzere temel elektronik tablo işlevselliği sağlar.

### ✨ Temel Özellikler

- **Izgara Tabanlı Veri Depolama**: 10 sütuna (A-J) ve 15 satıra (1-15) kadar destek
- **Dinamik Hücre Yönetimi**: Otomatik tür takibi ile hem metin hem de tam sayı değerleri saklar
- **Esnek İşlemler**: Veri manipülasyonu için 16 güçlü işlem
- **Akıllı Görüntüleme**: Uzun değerleri (>5 karakter) otomatik olarak "_" göstergesiyle kısaltır
- **Hata Yönetimi**: Açıklayıcı hata mesajlarıyla kapsamlı doğrulama
- **Veri Kalıcılığı**: Elektronik tablo içeriğini `spreadsheet.txt` dosyasına otomatik kaydeder

### 🚀 Başlangıç

#### Gereksinimler
- .NET SDK (C# destekleyen herhangi bir sürüm)
- Terminal/Komut İstemi

#### Kurulum

1. Depoyu klonlayın:
```bash
git clone https://github.com/AliYigitOzudogru/TUI-Based-Excel-Program.git
cd TUI-Based-Excel-Program
```

2. Programı derleyin:
```bash
csc Program.cs
```

3. Uygulamayı çalıştırın:
```bash
Program.exe
```

Veya tercih ettiğiniz IDE'yi kullanın (Visual Studio, Rider, VS Code).

### 📖 İşlem Kılavuzu

#### Temel Hücre İşlemleri

**1. AssignValue** - Hücreye veri ata
```
>> AssignValue(C4, integer, 45)
>> AssignValue(B9, string, MerhabaDunya)
```

**2. ClearCell** - Belirli hücreyi temizle
```
>> ClearCell(E8)
```

**3. ClearAll** - Tüm tabloyu temizle
```
>> ClearAll()
```

**Hücre Sorgulama** - Tam içeriği görüntüle
```
>> E8
```

#### Yapı İşlemleri

**4. AddRow** - Yeni satır ekle
```
>> AddRow(5, up)
>> AddRow(8, down)
```

**5. AddColumn** - Yeni sütun ekle
```
>> AddColumn(C, right)
>> AddColumn(D, left)
```

#### Kopyalama İşlemleri

**6. Copy** - Hücreyi başka hücreye kopyala
```
>> Copy(E8, A8)
```

**7. CopyColumn** - Tüm sütunu kopyala
```
>> CopyColumn(E, A)
```

**8. CopyRow** - Tüm satırı kopyala
```
>> CopyRow(3, 7)
```

#### Kesme İşlemleri

**9. X** - Hücreyi kes ve yapıştır
```
>> X(E8, A8)
```

**10. XColumn** - Sütunu kes ve yapıştır
```
>> XColumn(E, A)
```

**11. XRow** - Satırı kes ve yapıştır
```
>> XRow(3, 7)
```

#### Matematiksel & Metin İşlemleri

**12. Çarpma/Tekrar (*)** 
- Tam Sayı × Tam Sayı = Çarpma
- Metin × Tam Sayı = Tekrar (pozitif: normal, negatif: ters)
```
>> *(C4, A13, I14)        // 4 * 4 = 16
>> *(H14, H8, J1)         // "ABc" * 2 = "ABcABc"
>> *(I10, H14, A14)       // "ABc" * -2 = "cBAcBA"
```

**13. Toplama/Birleştirme (+)**
- Tam Sayı + Tam Sayı + Tam Sayı = Toplama
- Herhangi bir Metin varsa = Birleştirme (büyük/küçük harf dönüşümü ile)
```
>> +(A5, C4, D4)          // 6875 + 45 = 6920
>> +(E11, F11, G11, J6)   // Harf dönüşümü seçimi ile birleştirme
```

**14. Bölme/Metin Dilimleme (/)**
- Tam Sayı ÷ Tam Sayı = Bölme
- Metin ÷ Tam Sayı = İlk/son kısmı al
```
>> /(B9, H8, J1)          // "Algorithm" / 2 = "Algorith"
>> /(I10, B9, A14)        // "Algorithm" / -2 = "esruoc m"
```

**15. Çıkarma/Karakter Kaldırma (-)**
- Tam Sayı - Tam Sayı = Çıkarma
- Metin - Tam Sayı = ASCII karakterini kaldır
- Metin - Metin = Alt dizeyi kaldır
```
>> -(C4, I10, I14)        // 45 - (-2) = 47
>> -(G7, G6, J1)          // "CAATBA" - 65('A') = "CTB"
>> -(F3, D1, A14)         // "A3B4 CVB4E" - "B4" = "A3 CVE"
```

**16. Şifreleme (#)**
- Metin karakterlerinin ASCII değerlerini tam sayı kadar kaydırır
- Geçerli kaydırma aralığı: [-20, 30]
```
>> #(A1, A13, D5)         // "Apple" + 4 = "Ettpi"
>> #(H8, B9, J10)         // "Algorithm" + 2 = "Epksvmxlq"
```

### 🔧 Teknik Uygulama

#### Tasarım Özellikleri

- **Harici Ayrıştırma Kütüphanesi Yok**: Tüm projeде özel girdi doğrulama mantığı
- **2D Dizi Mimarisi**: Elektronik tablo ızgarası için ana veri yapısı
- **Manuel Tür Kontrolü**: TryParse kullanılmadan tüm işlemler için açık doğrulama
- **Prosedürel Yaklaşım**: Fonksiyonlar ve prosedürler aracılığıyla organize edilmiş
- **Karakter Bazında İşleme**: Manuel olarak uygulanan metin manipülasyonları

#### Veri Yapısı
```
- Ana Izgara: string[15,10]
- Tür Takipçisi: string[15,10] 
- Başlangıç Boyutu: 8 satır × 5 sütun
```

#### Hata Yönetimi
- Tüm işlemler için sınır doğrulama
- Tür uyumluluk kontrolleri
- Parametre sayısı doğrulama
- Her senaryo için özel hata mesajları

### 🛡️ Hata Mesajları

- `Operation is done!` - Başarılı
- `Illegal position assignment!` - Geçersiz hücre/satır/sütun referansı
- `Out of bounds exception!` - Maksimum ızgara boyutunu aşıyor
- `Illegal operation!` - Tür uyuşmazlığı veya geçersiz parametre kombinasyonu
- Belirli işlem kısıtlamaları için özel mesajlar

### 🎯 Kısıtlamalar

- Maksimum ızgara: 10 sütun × 15 satır
- (-) işlemi için ASCII karakter aralığı: [33, 126]
- Şifreleme kaydırma aralığı: [-20, 30]
- Tüm matematiksel işlemler atanmış hücreler gerektirir (atanmamış işlenen yok)

### 👨‍💻 Geliştirme Notları

Bu proje temel programlama kavramlarını vurgular:
- Manuel girdi ayrıştırma ve doğrulama
- Dizi manipülasyonu ve sınır yönetimi
- Yerleşik yardımcılar olmadan metin işleme
- Prosedürel ayrıştırma
- Hata öncelikli tasarım felsefesi

### 📜 Lisans

Bu proje eğitim ve kişisel kullanım için mevcuttur.

---

**Project Repository**: [GitHub](https://github.com/AliYigitOzudogru/TUI-Based-Excel-Program)

**Created with**: C# | .NET | Console Application

**Last Updated**: January 2026
