/**
 * ============================================================
 *  КОНФИГ ЦЕН И ЛИМИТОВ (для ИП)
 * ------------------------------------------------------------
 *  Единственный файл, который нужно править для смены цен,
 *  лимитов, пакетов и текстов.
 * ============================================================
 */

window.PRICING = {

    /* 1. Бесплатный лимит: 20 операций, сброс через 30 дней */
    freeMonthlyLimit: 20,
    freeResetDays: 30,

    /* 2. Стоимость операций в баллах (1 балл = 1 операция) */
    costs: {
        compressPdf: 1,
        convertFile: 1,
        convertPdfPage: 1,
        generateBatch: 1,
        comparePdf: 3,
        compareRisks: 1
    },

    /* 3. Лимиты для защиты от злоупотреблений */
    limits: {
        // PDF (сжатие / разделение / маркировка / подпись)
        maxPdfFiles: 20,
        maxPdfPages: 1000,
        maxSplitParts: 50,

        // Конвертер форматов
        maxConvertFiles: 10,
        maxConvertFileSizeMB: 100,

        // Генератор документов из Excel
        maxExcelRows: 2000,
        maxGeneratedDocs: 2000,

         // Сравнение договоров
        maxCompareFileSizeMB: 20,
        maxComparePages: 100,
    },

    /* 4. Пакеты для покупки */
    packages: [
        { id: 'pack_50', credits: 50, price: 50, label: 'Стартовый' },
        { id: 'pack_100', credits: 100, price: 95, label: 'Базовый', bonus: 5 },
        { id: 'pack_250', credits: 250, price: 220, label: 'Оптимальный', bonus: 30 },
        { id: 'pack_500', credits: 500, price: 400, label: 'Выгодный', bonus: 100 }
    ],

    /* 5. Минимальная покупка */
    minPurchase: 50,

    /* 6. Валюта */
    currency: '₽',
    currencyCode: 'RUB',

    /* 7. Фискализация */
    fiscalization: {
        enabled: true,
        provider: 'yookassa',
        requireEmail: true,
        defaultEmail: 'noreply@list2doc.ru'
    },

    /* 8. Реквизиты ИП */
    legal: {
        entityName: 'ИП Приходько Сергей Сергеевич',
        inn: '701729214964',
        ogrnip: '323470400073210',
        offerUrl: '/offer.html',
        privacyUrl: '/privacy.html'
    },

    /* 9. Тексты для интерфейса */
    messages: {
        freeLimitReached:
            'Бесплатные 20 операций в этом месяце закончились. ' +
            'Купите пакет баллов — от 50 ₽.',
        notEnoughCredits:
            'Недостаточно баллов для этой операции. Пополните баланс.',
        purchaseSuccess:
            'Баллы зачислены на ваш счёт. Спасибо!',
        emailRequired:
            'Укажите email — на него придёт чек по требованию 54-ФЗ.'
    }
};