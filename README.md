# ⚡ Assignment #4: Forms and Validation (HTML + Bootstrap)

**Astana IT University**  
**Course:** Web Technologies (2nd Year, 1st Trimester)  
**Student:** Beisenbay Estai (Бейсенбай Естай)  
**Group:** IT-2511  
**GitHub Repository:** [https://github.com/spongebobkzkz/assignment_WT_4](https://github.com/spongebobkzkz/assignment_WT_4)  
**Live Demo:** [https://spongebobkzkz.github.io/assignment_WT_4/](https://spongebobkzkz.github.io/assignment_WT_4/)

---

## 📋 Жоба мазмұны / Project Overview

Бұл тапсырмада **HTML5 native validation**, **Bootstrap 5.3 form components**, **Usability & Accessibility (WCAG 2.1)** және мобильді бейімделу (Responsive Design) талаптары толық орындалды. Тапсырма ережесіне сәйкес:
- ❌ **JavaScript қолданылмаған** (No JS)
- ❌ **`novalidate` атрибуты қолданылмаған**
- ❌ **`pattern` атрибуты қолданылмаған** (Task 2-де тек таза HTML типтері мен шектеулері)
- ❌ **Absolute positioning және бекітілген өлшемдер (fixed widths) жоқ**

Жоба пайдаланушының бұрынғы **SPEED CARGO / SwiftCargo** стиліндегі ыңғайлы dark & gold навигациялық жүйесімен, сапалы карточкаларымен және модульдік құрылымымен безендірілген.

---

## 📁 Файлдар құрылымы / File Structure

```text
Assignment_4/
│
├── index.html                           # Портал / Басты шолу беті (Cargo Luxury UI)
├── task1.html                           # Task 1: Пойыз билетін брондау (Book a train ticket)
├── task2.html                           # Task 2: AITU Hack 2026 тіркелу (HTML5 Native Validation)
├── task3.html                           # Task 3: Дұрыс емес форманы түзету (Fix a bad form)
│
├── css/
│   └── custom.css                       # Montserrat қарпі, палитра (#0F6B5F, #F2A93B), :user-invalid стилі
│
├── screenshots/
│   ├── task1_1280px.png                 # Task 1 компьютердегі көрінісі (1280px)
│   ├── task1_375px.png                  # Task 1 телефондағы көрінісі (375px)
│   ├── task2_1280px.png                 # Task 2 тексеру кестесімен көрінісі
│   ├── task3_1280px.png                 # Task 3 түзетілген форма және талдау кестесі
│   └── index_1280px.png                 # Басты бет скриншоты
│
├── Assignment4_Beisenbay_Estai.docx     # Ресми есеп (Word форматы, скриншоттар мен кестелермен)
├── Assignment4_Beisenbay_Estai.pdf      # Ресми есеп (PDF форматы)
├── generate_report.ps1                  # Есепті автоматты түрде жасау скрипті (PowerShell + Word COM)
├── take_screenshots.js                  # Headless Edge арқылы дәл скриншот түсіру скрипті
└── README.md                            # Толық сипаттама және қорғауға (defense) дайындық
```

---

## 🎯 Тапсырмалардың толық орындалуы / Task Details

### 🚆 Task 1. Book a Train Ticket (Bootstrap Forms + Custom CSS)
- **Түстер палитрасы:** 
  - Негізгі жасыл: `#0F6B5F` (Header, Active "One way" батырмасы, `+7` префиксі)
  - Қызғылт-сары: `#F2A93B` ("Find trains" батырмасы)
  - Қара-жасыл мәтін: `#16302B`
  - Фон түсі: `#F4F7F5`
  - Жиек түсі: `#E6EFEC`
- **Қаріп:** Google Fonts — `'Montserrat', sans-serif`.
- **Компоненттер:**
  1. **One way / Round trip:** Екі радио-батырмадан (`<input type="radio" class="btn-check">`) жасалған секторлы ауыстырғыш (segmented toggle).
  2. **From & To:** Bootstrap-тің `.form-floating` компоненті (Astana / Almaty).
  3. **Departure, Return, Passengers:** 3 бағанды тор (`col-12 col-md-4`). Жолаушылар саны: 1–6 ересек адам (`1 adult` – `6 adults`).
  4. **Carriage class:** Үш инлайн радио-батырма (`Seat`, `Coupé`, `Sleeper`).
  5. **Phone & Email:** `.input-group` компоненті, жасыл `+7` префикс блогы, телефон және email өрісі.
  6. **SMS & Ереже:** `.form-switch` арқылы SMS жіберу және ережеге келісу чекбоксы.
  7. **Батырмалардың мобильді реті:** Телефонда (375px) HTML ретін өзгертпестен **"Find trains"** батырмасы **"Clear"**-дің үстінде орналасады (`d-flex flex-column-reverse flex-md-row`).
  8. **"Clear" батырмасы:** Форманы тазарту үшін `type="reset"` қолданылды (JS-сіз жұмыс істейді).
  9. **Қолжетімділік:** Барлық өрістер байланысқан `<label for="...">` белгілеріне ие, тек пернетақтамен толық басқарылады (`Tab`, `Space`, `Enter`).

---

### 💻 Task 2. Registration for AITU Hack 2026 (HTML5 Validation)
Форма тек браузердің тума (native) HTML атрибуттары арқылы тексеріледі:
- `pattern` атрибуты **мүлдем қолданылмаған**!
- JavaScript **жоқ**, `novalidate` **жоқ**!

| Өріс | Ереже | Қолданылған HTML атрибуты | Қабылдайды (Accept) | Қабылдамайды (Reject) |
| :--- | :--- | :--- | :--- | :--- |
| **Full name \*** | 2–40 таңба | `type="text" minlength="2" maxlength="40" required` | `Anna-Maria Lee` | `A` |
| **University email \*** | Жарамды email | `type="email" required` | `a.fedenko@astanait.edu.kz` | `a.fedenko`, `a.fedenko@` |
| **Age \*** | 16-дан 30-ға дейінгі бүтін сан | `type="number" min="16" max="30" step="1" required` | `18` | `15`, `31` |
| **Team size \*** | 1-ден 4-ке дейінгі бүтін сан | `type="number" min="1" max="4" step="1" required` | `4` | `5`, `2.5` |
| **Track \*** | Web, Mobile, AI, GameDev | `<select required>` + `<option value="" disabled selected hidden>Choose...</option>` | `Web` | `"Choose…"` |
| **Arrival day \*** | 20–22 Қараша 2026 | `type="date" min="2026-11-20" max="2026-11-22" required` | `21.11.2026` | `25.11.2026` |
| **Check-in time \*** | 09:00–18:00, әр 30 минут | `type="time" min="09:00" max="18:00" step="1800" required` *(1800 сек = 30 мин)* | `10:30` | `10:15`, `19:00` |
| **Entry fee \*** | 0–5000 ₸, 500 ₸ қадаммен | `type="number" min="0" max="5000" step="500" required` | `1500` | `1250`, `6000` |
| **Experience level \*** | 3 радио-батырманың бірі | `type="radio" name="experienceLevel" required` | `Beginner` | таңдалмаған |
| **GitHub profile** | Қосымша, толық сілтеме | `type="url"` *(required қойылмаған)* | `(бос)`, `https://github.com/yerassyl-dev` | `github.com/yerassyl-dev` |
| **Password \*** | 8–20 таңба | `type="password" minlength="8" maxlength="20" required` | `Astana2026` | `short` |
| **Motivation \*** | 50–300 таңба | `<textarea minlength="50" maxlength="300" required>` | `120 characters` | `49 characters` |
| **Rules agreement \***| Келісім чекбоксы | `type="checkbox" required` | белгіленген | белгіленбеген |

#### 💡 Заманауи CSS псевдо-кластары (`:user-invalid` & `:user-valid`):
Парақ ашылғанда қателер қызарып тұрмауы үшін тапсырма шартына сәйкес мынадай CSS ережесі жазылды:
```css
.task2-form input:user-invalid,
.task2-form select:user-invalid,
.task2-form textarea:user-invalid {
  border-color: #DC2626 !important;
}

.task2-form input:user-invalid ~ .invalid-feedback,
.task2-form select:user-invalid ~ .invalid-feedback,
.task2-form textarea:user-invalid ~ .invalid-feedback {
  display: block !important;
}

.task2-form input:user-valid:not([type="checkbox"]):not([type="radio"]),
.task2-form select:user-valid,
.task2-form textarea:user-valid {
  border-color: #16A34A !important;
}
```
Нәтижесінде, қолданушы өріспен жұмыс жасап, қате енгізгеннен кейін ғана қызыл жиек пен ескерту мәтіні шығады!

---

### 🛠️ Task 3. Fix a Bad Form (HTML + Bootstrap)
Берілген кафенің үстел брондау формасындағы **12 өрескел қателік** табылып, талданды және түзетілді:

1. **`Your name` өрісі:** `<label>` және `name` атрибуты жоқ болған. `id="resName" name="customerName" required autocomplete="name"` қосылды.
2. **`Email` өрісі:** `type="text"` деп қате берілген. Ол мобильді телефонда `@` пернесін шығармайды. `type="email"` деп түзетілді.
3. **`Phone` өрісі:** `<label>Phone</label>` тегінде `for` атрибуты болмағандықтан, мәтінді басқанда инпут белсенді болмайтын. 
   - **Біз таңдаған шешім:** `<label for="resPhone">` және `<input type="tel" id="resPhone">` жасау.
   - **Балама шешім (Alternative fix):** Инпутты тікелей label ішіне орау: `<label>Phone <input type="tel"></label>`.
   - **Неліктен біз өз шешімімізді таңдадық:** Bootstrap 5 grid жүйесі мен `.form-label` стилі бөлек тегтер болғанда таза және дұрыс жұмыс істейді.
4. **`Guests` өрісіндегі ID сәйкессіздігі (ID mismatch):** Белгіде `for="guests"` болса, инпутта `id="people"`. Түзетілді: `id="resGuests"`, `type="number" min="1" max="30" step="1"`.
5. **`Date` өрісі:** Жай мәтіндік өріс болған және `id="date"` төмендегі радиомен қайталанған. `type="date"`, бірегей `id="resDate"` қойылды.
6. **`Inside` және `Terrace` радиолары:** `name="inside"` және `name="terrace"` деп екі түрлі аталған, соның кесірінен екеуі бірдей таңдалып қалатын. Оған қоса `id="date"` қайталанған және мәтін үшін `<label>` жоқ. `name="seatingPreference"` деп біріктіріліп, бірегей ID және label берілді.
7. **`Seating:` параграфы:** Семантикалық емес `<p>` орнына экран оқу құралдары (screen readers) үшін `<fieldset>` және `<legend>` қолданылды.
8. **`<select>` селекторы:** Атауы (`name`) және label жоқ. Түзетілді: `name="mealSlot"`, `Lunch` және `Dinner` таңдауларына нақты `value` берілді.
9. **`<textarea>` пікірлер:** Белгісі, ID-і және серверге жіберетін `name="customerComments"` қосылды.
10. **`I agree to the terms` чекбоксы:** Label байланыспаған, телефоннан кішкентай 13px шаршыны басу өте қиын болған. Bootstrap `.form-check` құрылымына көшірілді.
11. **`<div class="btn btn-primary">Reserve</div>` (ЕҢ ӨРЕСКЕЛ ҚАТЕ):**
    - `<div>` батырма емес! Оған пернетақтамен (`Tab`) өту мүмкін емес, `Enter` басуға болмайды және ең бастысы — ол форманы серверге ЖІБЕРМЕЙДІ (never submits the form)!
    - **Біз таңдаған шешім:** Нағыз `<button type="submit" class="btn btn-primary">Reserve Table</button>` тегіне ауыстыру.
    - **Балама шешім (Alternative fix):** `<input type="submit" value="Reserve Table" class="btn btn-primary">` қолдану.
    - **Неліктен біз өз шешімімізді таңдадық:** `<button>` тегі икемдірек, оның ішіне белгішелер (иконкалар) қосуға болады және ол қазіргі заманғы Bootstrap стандартына сай.
12. **`<form>` тегінде атрибуттардың болмауы:** `action="#" method="post"` қосылды.

Форма **W3C Markup Validator** (World Wide Web Consortium) стандартынан қатесіз (0 errors, 0 warnings) өтеді.

---

## 🎓 Қорғауға (Defense) дайындық сұрақ-жауаптары (50 ұпай)

Мұғалім қорғауда сұрауы мүмкін негізгі сұрақтар мен дайын жауаптар:

1. **Сұрақ:** Task 1-де телефонда "Find trains" қалай "Clear"-ден жоғары тұр, бірақ HTML-де Clear бірінші жазылған?
   - **Жауап:** Бұл Bootstrap-тің flexbox утилитасы арқылы жасалды: `d-flex flex-column-reverse flex-md-row`. Телефонда (`< 768px`) `flex-column-reverse` элементтердің визуальды бағытын кері бұрады, сондықтан екінші тұрған "Find trains" үстіге шығады. Ал компьютерде (`>= 768px`) `flex-md-row` оларды көлденең қалыпты ретпен орналастырады.
2. **Сұрақ:** Неге Task 2-де бетті ашқанда қателер қызарып кетпейді?
   - **Жауап:** Біз заманауи CSS псевдо-класы `:user-invalid`-ті пайдаландық. Стандартты `:invalid` псевдо-класы бет ашылған бойда бос өрістерді бірден қызартып көрсетеді, ал `:user-invalid` тек пайдаланушы өрісті өзгерткеннен кейін ғана қате болса іске қосылады.
3. **Сұрақ:** Уақыт өрісінде әр 30 минут қалай тексеріледі?
   - **Жауап:** HTML5 `type="time"` өрісінде `step` атрибуты **секундпен** өлшенеді! 30 минут = 30 × 60 = 1800 секунд (`step="1800"`). Сондықтан 10:30 қабылданады, ал 10:15 (15 минуттық ауытқу) қабылданбайды.
4. **Сұрақ:** Task 3-тегі `<div class="btn">` неліктен үлкен қате?
   - **Жауап:** `<div>` семантикалық элемент емес. Ол `tabindex`-сіз болғандықтан пернетақта фокусын қабылдамайды, screen reader оны батырма ретінде оқымайды және оны басқанда браузер форманың `submit` оқиғасын тудырмайды, нәтижесінде деректер серверге ешқашан жетпейді.

---

## 🚀 Жобаны жергілікті жерде ашу (How to Run)

Кез-келген браузерде (Chrome, Edge, Firefox) мына файлдардың кез-келгенін тікелей екі рет шертіп ашуға болады:
1. `index.html` — Басты портал
2. `task1.html` — Task 1 формасы
3. `task2.html` — Task 2 формасы
4. `task3.html` — Task 3 формасы

Сондай-ақ ресми есеп Word (`Assignment4_Beisenbay_Estai.docx`) және PDF (`Assignment4_Beisenbay_Estai.pdf`) форматында папка ішінде толық дайын!
