# Руководство по регистрации доменов

**Домены:** `pengboaviation.com` · `pengbo.aero`  
**Проверка доступности:** 14 июня 2026 — оба домена **свободны** (RDAP 404)  
**Статус:** ожидает оплаты и действий владельца

> Регистрация требует банковской карты и личного кабинета у регистратора.  
> Эти шаги выполняются вручную; ниже — готовые данные для копирования в формы.

---

## Порядок действий (кратко)

| Шаг | Действие | Срок |
|-----|----------|------|
| **1** | Зарегистрировать **pengboaviation.com** | Сегодня, ~10 мин |
| **2** | Подать заявку на **SITA Membership ID** для .aero | Сегодня, ~15 мин |
| **3** | Дождаться письма от SITA с ID и Auth Code | 1–2 рабочих дня |
| **4** | Зарегистрировать **pengbo.aero** у аккредитованного регистратора | После шага 3 |
| **5** | Настроить DNS (рекомендуется Cloudflare) | После шагов 1 и 4 |

---

## Шаг 1. Регистрация pengboaviation.com

### Рекомендуемый регистратор: Cloudflare Registrar

- Цена без наценки (~$10–11/год)
- Бесплатный DNS, CDN, SSL
- Ссылка: https://dash.cloudflare.com/sign-up → **Domain Registration** → поиск `pengboaviation.com`

**Альтернативы:** [Porkbun](https://porkbun.com) · [Namecheap](https://www.namecheap.com)

### Данные registrant (скопировать в форму)

| Поле | Значение |
|------|----------|
| Organization | Jixi Huijie Aircraft Research and Development Center Co., Ltd. |
| Registrant name | Geng Huijie |
| Address line 1 | No. 11, Shop, Building 1, Baile Jiayuan, Xingguo Road |
| City | Jixi |
| State/Province | Heilongjiang |
| Postal code | 158100 |
| Country | China (CN) |
| Phone | +86.13704631315 |
| Email | jixi_huijie@proton.me |

**Китайский адрес (если запрашивают):**  
黑龙江省鸡西市鸡冠区兴国路百乐家园1号楼11号门市

### После оплаты

1. Включить **Auto-renew** (автопродление).
2. Включить **Registrar Lock** / Transfer Lock.
3. Записать дату истечения в календарь.
4. *(Опционально)* Перенести DNS-зону в Cloudflare, если регистрация была у другого регистратора.

---

## Шаг 2. Заявка на SITA Membership ID (обязательно для .aero)

Форма: https://information.aero/registration/aero_id_authenticate

Проверка имени: https://information.aero/ — поле **Check name** → `pengbo`

### Рекомендуемые значения формы

| Поле формы SITA | Значение |
|-----------------|----------|
| **Registrant Group** | `Aviation Supplier or Service Provider` |
| **Organization** | `Jixi Huijie Aircraft Research and Development Center Co., Ltd.` |
| **Credential Type** | `Other (2Z)` — или `Business registration` / `Company website`, если доступно |
| **Credential Value** | `USCC 91230300MA1CPEPW5B; CASC aero.cascpooling.com registrant; GA aircraft operator (3 aircraft)` |

### Контактные данные

| Поле | Значение |
|------|----------|
| Contact name | Geng Huijie |
| Email | jixi_huijie@proton.me |
| Phone | +86 137 0463 1315 |
| Address | No. 11, Shop, Building 1, Baile Jiayuan, Xingguo Road, Jiguan District |
| City | Jixi |
| State/Province | Heilongjiang |
| Country | China |
| Postal code | 158100 |

### Документы для подтверждения (приложить по email, если SITA запросит)

1. Скан **营业执照** (Business License) с USCC `91230300MA1CPEPW5B`
2. Профиль компании: `ji xi hui jie/base/company-profile-en.md` (экспорт в PDF)
3. Скриншот / подтверждение регистрации на **CASC** (aero.cascpooling.com)
4. Краткое описание деятельности (1 абзац):

```
Jixi Huijie Aircraft Research and Development Center Co., Ltd. (est. August 2021)
operates general aviation aircraft (3 aircraft), conducts agricultural and forestry
aviation services, aircraft R&D, helicopter development projects, and international
supply of aviation components (locks, fasteners, GSE) through partner manufacturers.
Registered on CASC Aviation Materials platform. USCC: 91230300MA1CPEPW5B.
```

### Ожидаемый результат

SITA пришлёт на email:

- **Membership ID** (Aero ID)
- **Auth Code** (пароль)

Срок: обычно **1–2 рабочих дня**. Контакт SITA: aero.enquiries@sita.aero

---

## Шаг 3. Регистрация pengbo.aero

### Регистратор: 101domain (аккредитован для .aero)

- Страница: https://www.101domain.com/aero.htm
- Поиск: `pengbo.aero`
- Цена: ~$68/год (регистрация), ~$90/год (продление)

### При оформлении заказа

Ввести полученные от SITA:

| Поле | Значение |
|------|----------|
| Aero ID / Membership ID | *(из письма SITA)* |
| Auth Code / Aero password | *(из письма SITA)* |

Registrant data — те же, что в шаге 1.

> **Важно:** один Membership ID можно использовать для нескольких доменов .aero, если понадобятся защитные имена (pbt.aero и т.д.).

---

## Шаг 4. DNS (после регистрации обоих доменов)

### Рекомендуемая схема

```
pengboaviation.com     → основной сайт
www.pengboaviation.com → CNAME на pengboaviation.com
pengbo.aero            → 301 redirect на https://www.pengboaviation.com
www.pengbo.aero        → 301 redirect на https://www.pengboaviation.com
```

### Записи для почты (когда появится хостинг)

| Тип | Имя | Значение |
|-----|-----|----------|
| MX | @ | *(по инструкции почтового провайдера)* |
| TXT | @ | `v=spf1 ...` |
| TXT | _dmarc | `v=DMARC1; p=none; rua=mailto:jixi_huijie@proton.me` |

Пока сайта нет — достаточно parking-страницы или временной A-записи регистратора.

---

## Ориентировочный бюджет

| Позиция | Стоимость |
|---------|-----------|
| pengboaviation.com (1 год) | ~$10–15 |
| pengbo.aero (1 год) | ~$68–99 |
| **Итого первый год** | **~$78–114** |

---

## Чеклист

- [ ] pengboaviation.com зарегистрирован
- [ ] Auto-renew включён для .com
- [ ] Заявка SITA Membership ID отправлена
- [ ] Получены Membership ID и Auth Code
- [ ] pengbo.aero зарегистрирован
- [ ] pengbo.aero → редирект на pengboaviation.com
- [ ] WHOIS-контакты актуальны
- [ ] Дата продления внесена в календарь

---

## Если домен уже заняли

Пока вы регистрируетесь, возможен cybersquatting. Действия:

1. Немедленно проверить WHOIS на сайте регистратора.
2. Запасные варианты: `pengbo-aviation.com`, `pbtaviation.com`.
3. Для .aero: `pengbo-aviation.aero` (если `pengbo` зарезервировано).

---

*Связанные документы: [domain-naming-report-ru.md](./domain-naming-report-ru.md)*
