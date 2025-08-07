document.addEventListener('DOMContentLoaded', () => {

    // --- 1. DATA: ALL LESSONS, TOPICS, AND FUNCTIONS ---
    const content = {
        // --- Static text for the header ---
        header: {
            main_title: { en: 'Microsoft Excel Training Syllabus', az: 'Microsoft Excel Tədris Proqramı' },
            prepared_by: { en: 'Prepared by: Nizami Tahirov | Interactive Version', az: 'Hazırladı: Nizami Tahirov | İnteraktiv Versiya' }
        },
        // --- Full lesson data ---
        lessons: {
            'l1': {
                title: { en: 'Lesson 1: Introduction & Excel Basics', az: 'Dərs 1: Giriş və Excel-in Əsasları' },
                topics: [
                    { title: { en: 'General Overview of the Excel Interface', az: 'Excel İnterfeysinə Ümumi Baxış' }, desc: { en: 'Get familiar with the main components: Ribbon, Quick Access Toolbar, Formula Bar, and Status Bar.', az: 'Əsas komponentlər ilə tanışlıq: Lent (Ribbon), Cəld Müraciət Paneli, Formula Zolağı və Status Zolağı.' }},
                    { title: { en: 'Detailed Look at the Excel Interface', az: 'Excel İnterfeysinə Detallı Baxış' }, desc: { en: 'A deeper dive into the different tabs (Home, Insert, Page Layout, etc.) and their core functionalities.', az: 'Fərqli tabların (Home, Insert, Page Layout və s.) və onların əsas funksionallıqlarının dərindən öyrənilməsi.' }},
                    { title: { en: 'Navigation and Selections', az: 'Naviqasiya və Seçimlər' }, desc: { en: 'Learn efficient ways to move around your worksheet and select cells, ranges, rows, and columns.', az: 'İş vərəqində səmərəli hərəkət etmək və xanaları, diapazonları, sətirləri və sütunları seçmək.' }},
                    { title: { en: 'View Options', az: 'Görünüş Seçimləri' }, desc: { en: 'Customize your workspace with different views like Normal, Page Layout, and Page Break Preview. Master zoom and freezing panes.', az: 'İş sahənizi Normal, Səhifə Düzəni və Səhifə Sonu Kəsimi kimi fərqli görünüşlərlə fərdiləşdirin.' }},
                    { title: { en: 'Data Types', az: 'Data Tipləri' }, desc: { en: 'Understand the difference between text, numbers, dates, and logical values.', az: 'Mətn, ədəd, tarix və məntiqi dəyərlər arasındakı fərqi anlamaq.' }},
                    { title: { en: 'Cell Formats', az: 'Xana Formatları' }, desc: { en: 'Apply basic formatting to make your data readable and professional.', az: 'Məlumatlarınızı oxunaqlı və peşəkar etmək üçün əsas formatlaşdırma tətbiq etmək.' }},
                    { title: { en: 'Data Entry, Editing, and Deletion', az: 'Məlumatların Daxil Edilməsi, Redaktəsi və Silinməsi' }, desc: { en: 'The fundamentals of inputting information, making changes, and clearing cell contents.', az: 'Məlumat daxil etmə, dəyişiklik etmə və xana məzmununu təmizləmənin əsasları.' }},
                    { title: { en: 'Fill Handle', az: 'Doldurma Tutacağı («Fill Handle»)' }, desc: { en: 'Discover the power of the Fill Handle to quickly copy data and create series of numbers or dates.', az: 'Məlumatları sürətlə kopyalamaq və ədəd və ya tarix seriyaları yaratmaq üçün Doldurma Tutacağının gücünü kəşf edin.' }},
                    { title: { en: 'Flash Fill', az: 'Ani Doldurma (Flash Fill)' }, desc: { en: 'Let Excel detect patterns and automatically fill data for you, saving massive amounts of time.', az: 'Excel-in nümunələri aşkar etməsinə və sizin üçün məlumatları avtomatik doldurmasına imkan verin, bu da xeyli vaxta qənaət edir.' }}
                ]
            },
            'l2': {
                title: { en: 'Lesson 2: Advanced Formatting & Data Handling', az: 'Dərs 2: Təkmil Formatlama və Data İdarəetməsi' },
                topics: [
                    { title: { en: 'Cell Formats & Custom Formats', az: 'Xana Formatları və Fərdi Formatlar' }, desc: { en: 'Go beyond the basics to master Number, Currency, Date formats, and create your own unique formats.', az: 'Əsaslardan kənara çıxaraq Ədəd, Valyuta, Tarix formatlarını mənimsəyin və öz unikal formatlarınızı yaradın.' }},
                    { title: { en: 'Conditional Formatting', az: 'Şərtli Formatlama' }, desc: { en: 'Automatically highlight key data, trends, and exceptions using rules, data bars, color scales, and icon sets.', az: 'Qaydalar, data zolaqları, rəng şkalaları və ikon dəstlərindən istifadə edərək əsas məlumatları avtomatik olaraq vurğulayın.' }},
                    { title: { en: 'Custom Lists', az: 'Fərdi Siyahılar' }, desc: { en: 'Create and use your own custom sort and fill lists (e.g., product names, regions) for faster data entry.', az: 'Daha sürətli data girişi üçün öz fərdi çeşidləmə və doldurma siyahılarınızı (məsələn, məhsul adları, regionlar) yaradın.' }},
                    { title: { en: 'Cell References (Relative, Absolute, Mixed)', az: 'Xana İstinadları (Nisbi, Mütləq, Qarışıq)' }, desc: { en: 'Understand the critical difference between A1, $A$1, A$1, and $A1 to build flexible formulas.', az: 'Çevik düsturlar qurmaq üçün A1, $A$1, A$1 və $A1 arasındakı kritik fərqi anlayın.' }},
                    { title: { en: 'Copy & Paste Options (Paste Special)', az: 'Kopyalama və Yapışdırma Seçimləri (Xüsusi Yapışdırma)' }, desc: { en: 'Unleash the power of Paste Special to transpose data or paste only values, formats, or formulas.', az: 'Datanı transpozisiya etmək və ya yalnız dəyərləri, formatları və ya formulları yapışdırmaq üçün "Paste Special"-in gücündən istifadə edin.' }}
                ]
            },
            'l3': {
                title: { en: 'Lesson 3: Introduction to Formulas & Functions', az: 'Dərs 3: Formullara və Funksiyalara Giriş' },
                topics: [
                    { title: { en: 'Understanding Formulas vs. Functions', az: 'Formul və Funksiya Anlayışları' }, desc: { en: 'Learn the core difference between a manual calculation (formula) and a pre-built Excel command (function).', az: 'Əl ilə edilən hesablama (formul) ilə Excel-in hazır əmri (funksiya) arasındakı əsas fərqi öyrənin.' }},
                    { title: { en: 'Simple Mathematical Operations', az: 'Sadə Riyazi Əməliyyatlar' }, desc: { en: 'Perform basic arithmetic (+, -, *, /) and understand the order of operations.', az: 'Əsas riyazi əməliyyatları (+, -, *, /) yerinə yetirin və əməliyyatlar ardıcıllığını anlayın.' }},
                    { title: { en: 'Basic Mathematical Functions', az: 'Əsas Riyazi Funksiyalar' }, desc: { en: 'Get introduced to foundational math functions.', az: 'Əsas riyazi funksiyalarla tanış olun.' }}
                ]
            },
            'l4': {
                title: { en: 'Lesson 4: Statistical & Logical Functions', az: 'Dərs 4: Statistik və Məntiqi Funksiyalar' },
                topics: [
                    { title: {en: 'Simple Statistical Functions', az: 'Sadə Statistik Funksiyalar' }, desc: { en: 'Functions for basic statistical analysis.', az: 'Əsas statistik təhlil üçün funksiyalar.' },
                        syntax: [
                            { code: 'COUNT(value1, [value2], ...)', desc: {en: 'Counts cells containing numbers.', az: 'Tərkibində ədəd olan xanaları sayır.'} },
                            { code: 'COUNTA(value1, [value2], ...)', desc: {en: 'Counts non-empty cells.', az: 'Boş olmayan xanaları sayır.'} },
                            { code: 'COUNTBLANK(range)', desc: {en: 'Counts empty cells in a range.', az: 'Diapazondakı boş xanaları sayır.'} },
                            { code: 'AVERAGE(number1, [number2], ...)', desc: {en: 'Calculates the arithmetic mean.', az: 'Ədədi ortanı hesablayır.'} },
                            { code: 'MAX(number1, [number2], ...)', desc: {en: 'Finds the largest value.', az: 'Ən böyük dəyəri tapır.'} },
                            { code: 'MIN(number1, [number2], ...)', desc: {en: 'Finds the smallest value.', az: 'Ən kiçik dəyəri tapır.'} },
                            { code: 'MEDIAN(number1, [number2], ...)', desc: {en: 'Finds the middle value (median).', az: 'Orta dəyəri (median) tapır.'} },
                            { code: 'SMALL(array, k)', desc: {en: 'Finds the k-th smallest value.', az: 'k-ıncı ən kiçik dəyəri tapır.'} },
                            { code: 'LARGE(array, k)', desc: {en: 'Finds the k-th largest value.', az: 'k-ıncı ən böyük dəyəri tapır.'} }
                        ]
                    },
                    { title: { en: 'Logical Functions', az: 'Məntiqi Funksiyalar' }, desc: { en: 'Functions for making decisions.', az: 'Qərar vermək üçün funksiyalar.'},
                        syntax: [
                            { code: 'IF(logical_test, value_if_true, [value_if_false])', desc: {en: 'Performs a logical test.', az: 'Məntiqi bir test həyata keçirir.'} },
                            { code: 'IFERROR(value, value_if_error)', desc: {en: 'Returns a custom value if a formula results in an error.', az: 'Formula xəta verərsə, xüsusi dəyər qaytarır.'} },
                            { code: 'AND(logical1, [logical2], ...)', desc: {en: 'Checks if ALL conditions are TRUE.', az: 'BÜTÜN şərtlərin doğru olub olmadığını yoxlayır.'} },
                            { code: 'OR(logical1, [logical2], ...)', desc: {en: 'Checks if ANY condition is TRUE.', az: 'HƏR HANSI bir şərtin doğru olub olmadığını yoxlayır.'} },
                            { code: 'NOT(logical)', desc: {en: 'Reverses the logical value.', az: 'Məntiqi dəyəri tərsinə çevirir.'} }
                        ]
                    }
                ]
            },
            'l5': {
                title: { en: 'Lesson 5: Text Functions', az: 'Dərs 5: Mətn Funksiyaları' },
                topics: [
                    { title: {en: 'Text Manipulation Functions', az: 'Mətn Manipulyasiya Funksiyaları'}, desc: { en: 'A comprehensive set of functions for working with text strings.', az: 'Mətn sətirləri ilə işləmək üçün əhatəli funksiyalar toplusu.' },
                        syntax: [
                            { code: 'LEN(text)', desc: { en: 'Returns the number of characters in a text string.', az: 'Mətn sətrindəki simvolların sayını qaytarır.' }},
                            { code: 'UPPER(text)', desc: { en: 'Converts text to uppercase.', az: 'Mətni böyük hərflərə çevirir.' }},
                            { code: 'LOWER(text)', desc: { en: 'Converts text to lowercase.', az: 'Mətni kiçik hərflərə çevirir.' }},
                            { code: 'PROPER(text)', desc: { en: 'Capitalizes the first letter in each word.', az: 'Hər sözün ilk hərfini böyük edir.' }},
                            { code: 'LEFT(text, [num_chars])', desc: { en: 'Returns characters from the beginning of a text string.', az: 'Mətnin əvvəlindən simvolları qaytarır.' }},
                            { code: 'RIGHT(text, [num_chars])', desc: { en: 'Returns characters from the end of a text string.', az: 'Mətnin sonundan simvolları qaytarır.' }},
                            { code: 'MID(text, start_num, num_chars)', desc: { en: 'Returns characters from the middle of a text string.', az: 'Mətnin ortasından simvolları qaytarır.' }},
                            { code: 'TRIM(text)', desc: { en: 'Removes extra spaces from text.', az: 'Mətndən artıq boşluqları silir.' }},
                            { code: 'CONCAT(text1, [text2], ...)', desc: { en: 'Joins several text items into one text item.', az: 'Bir neçə mətni bir mətnə birləşdirir.' }},
                            { code: 'TEXTJOIN(delimiter, ignore_empty, text1, ...)', desc: { en: 'Joins text with a delimiter.', az: 'Mətnləri ayırıcı ilə birləşdirir.' }},
                            { code: 'FIND(find_text, within_text, [start_num])', desc: { en: 'Finds one text string within another (case-sensitive).', az: 'Bir mətni digərinin içində tapır (hərf həssasdır).' }},
                            { code: 'SEARCH(find_text, within_text, [start_num])', desc: { en: 'Finds one text string within another (not case-sensitive).', az: 'Bir mətni digərinin içində tapır (hərf həssas deyil).' }},
                            { code: 'REPLACE(old_text, start_num, num_chars, new_text)', desc: { en: 'Replaces characters within text based on position.', az: 'Mövqeyə əsaslanaraq mətndəki simvolları əvəz edir.' }},
                            { code: 'SUBSTITUTE(text, old_text, new_text, [instance_num])', desc: { en: 'Replaces existing text with new text in a string.', az: 'Mətndə mövcud mətni yeni mətnlə əvəz edir.' }},
                            { code: 'REPT(text, number_times)', desc: { en: 'Repeats text a given number of times.', az: 'Mətni verilmiş sayda təkrarlayır.' }},
                            { code: 'TEXT(value, format_text)', desc: { en: 'Formats a number and converts it to text.', az: 'Ədədi formatlayır və mətnə çevirir.' }}
                        ]
                    }
                ]
            },
             'l6': {
                title: { en: 'Lesson 6: Date & Time Functions', az: 'Dərs 6: Tarix və Zaman Funksiyaları' },
                topics: [
                     { title: {en: 'Date and Time Functions', az: 'Tarix və Zaman Funksiyaları'}, desc: { en: 'Functions for performing calculations with dates and times.', az: 'Tarix və zamanla hesablama aparmaq üçün funksiyalar.' },
                        syntax: [
                            { code: 'TODAY()', desc: {en: 'Returns the current date.', az: 'Cari tarixi qaytarır.'}},
                            { code: 'NOW()', desc: {en: 'Returns the current date and time.', az: 'Cari tarixi və zamanı qaytarır.'}},
                            { code: 'DAY(serial_number)', desc: {en: 'Returns the day of the month (1-31).', az: 'Ayın gününü qaytarır (1-31).'}},
                            { code: 'MONTH(serial_number)', desc: {en: 'Returns the month (1-12).', az: 'Ayı qaytarır (1-12).'}},
                            { code: 'YEAR(serial_number)', desc: {en: 'Returns the year.', az: 'İli qaytarır.'}},
                            { code: 'DATE(year, month, day)', desc: {en: 'Returns the serial number for a specific date.', az: 'Müəyyən bir tarix üçün seriya nömrəsini qaytarır.'}},
                            { code: 'DATEVALUE(date_text)', desc: {en: 'Converts a date in text format to a serial number.', az: 'Mətn formatında olan tarixi seriya nömrəsinə çevirir.'}},
                            { code: 'EDATE(start_date, months)', desc: {en: 'Returns a date that is a number of months before or after a start date.', az: 'Başlanğıc tarixindən müəyyən ay əvvəlki və ya sonrakı tarixi qaytarır.'}},
                            { code: 'EOMONTH(start_date, months)', desc: {en: 'Returns the last day of the month before or after a specified number of months.', az: 'Müəyyən sayda aydan əvvəlki və ya sonrakı ayın son gününü qaytarır.'}},
                            { code: 'WORKDAY(start_date, days, [holidays])', desc: {en: 'Returns a date that is a number of working days before or after a date.', az: 'Bir tarixdən müəyyən sayda iş günü əvvəl və ya sonra olan tarixi qaytarır.'}},
                            { code: 'NETWORKDAYS(start_date, end_date, [holidays])', desc: {en: 'Returns the number of whole working days between two dates.', az: 'İki tarix arasındakı tam iş günlərinin sayını qaytarır.'}},
                            { code: 'DATEDIF(start_date, end_date, unit)', desc: {en: 'Calculates the number of days, months, or years between two dates. (Unit: "D", "M", "Y").', az: 'İki tarix arasındakı günlərin, ayların və ya illərin sayını hesablayır. (Vahid: "D", "M", "Y").'}}
                        ]
                    }
                ]
            },
            'l7': {
                title: { en: 'Lesson 7: Conditional Aggregate Functions', az: 'Dərs 7: Şərti Riyazi Funksiyalar' },
                topics: [
                    { title: {en: 'Conditional Math Functions', az: 'Şərti Riyazi Funksiyalar'}, desc: { en: 'Functions that perform calculations based on conditions.', az: 'Şərtlərə əsasən hesablama aparan funksiyalar.' },
                        syntax: [
                            { code: 'SUMIF(range, criteria, [sum_range])', desc: { en: 'Sums cells that meet a single criterion.', az: 'Tək bir şərtə cavab verən xanaları cəmləyir.'}},
                            { code: 'COUNTIF(range, criteria)', desc: { en: 'Counts cells that meet a single criterion.', az: 'Tək bir şərtə cavab verən xanaları sayır.'}},
                            { code: 'AVERAGEIF(range, criteria, [average_range])', desc: { en: 'Averages cells that meet a single criterion.', az: 'Tək bir şərtə cavab verən xanaların ortalamasını tapır.'}},
                            { code: 'SUMIFS(sum_range, criteria_range1, criteria1, ...)', desc: { en: 'Sums cells that meet multiple criteria.', az: 'Çoxsaylı şərtlərə cavab verən xanaları cəmləyir.'}},
                            { code: 'COUNTIFS(criteria_range1, criteria1, ...)', desc: { en: 'Counts cells that meet multiple criteria.', az: 'Çoxsaylı şərtlərə cavab verən xanaları sayır.'}},
                            { code: 'AVERAGEIFS(average_range, criteria_range1, criteria1, ...)', desc: { en: 'Averages cells that meet multiple criteria.', az: 'Çoxsaylı şərtlərə cavab verən xanaların ortalamasını tapır.'}},
                            { code: 'MAXIFS(max_range, criteria_range1, criteria1, ...)', desc: { en: 'Finds the maximum value based on multiple criteria.', az: 'Çoxsaylı şərtlərə əsasən maksimum dəyəri tapır.'}},
                            { code: 'MINIFS(min_range, criteria_range1, criteria1, ...)', desc: { en: 'Finds the minimum value based on multiple criteria.', az: 'Çoxsaylı şərtlərə əsasən minimum dəyəri tapır.'}}
                        ]
                    },
                    { title: {en: 'Handling Duplicates', az: 'Dublikatların İdarə Edilməsi'}, desc: { en: 'Tools to find and remove duplicate values.', az: 'Təkrarlanan dəyərləri tapmaq və silmək üçün alətlər.' },
                        syntax: [
                             { code: 'UNIQUE(array, [by_col], [exactly_once])', desc: { en: 'Returns a list of unique values from a range (Office 365+).', az: 'Bir diapazondan unikal dəyərlərin siyahısını qaytarır (Office 365+).'}},
                             { code: 'Data > Remove Duplicates', desc: { en: 'A command on the Data tab to permanently delete duplicate rows.', az: 'Təkrarlanan sətirləri həmişəlik silmək üçün Data tabında olan bir əmr.'}}
                        ]
                    }
                ]
            },
            'l8': {
                title: { en: 'Lesson 8: Lookup & Reference Functions', az: 'Dərs 8: Axtarış və İstinad Funksiyaları' },
                topics: [
                     { title: {en: 'Lookup & Reference Functions', az: 'Axtarış və İstinad Funksiyaları'}, desc: { en: 'Powerful functions for finding and retrieving data.', az: 'Məlumatları tapmaq və gətirmək üçün güclü funksiyalar.' },
                        syntax: [
                            { code: 'VLOOKUP(lookup_value, table_array, col_index_num, ...)', desc: { en: 'Searches for a value in the first column of a table.', az: 'Cədvəlin ilk sütununda bir dəyər axtarır.'}},
                            { code: 'HLOOKUP(lookup_value, table_array, row_index_num, ...)', desc: { en: 'Searches for a value in the top row of a table.', az: 'Cədvəlin yuxarı sətrində bir dəyər axtarır.'}},
                            { code: 'XLOOKUP(lookup_value, lookup_array, return_array, ...)', desc: { en: 'The modern and flexible successor to VLOOKUP/HLOOKUP (O365+).', az: 'VLOOKUP/HLOOKUP-un müasir və çevik varisi (O365+).'}},
                            { code: 'INDEX(array, row_num, [column_num])', desc: { en: 'Returns a value from within a table or range based on position.', az: 'Mövqeyə əsasən cədvəl və ya diapazondan dəyər qaytarır.'}},
                            { code: 'MATCH(lookup_value, lookup_array, [match_type])', desc: { en: 'Finds the position of an item in a range.', az: 'Bir diapazonda elementin mövqeyini tapır.'}},
                            { code: 'CHOOSE(index_num, value1, [value2], ...)', desc: { en: 'Selects a value from a list of values.', az: 'Dəyərlər siyahısından bir dəyər seçir.'}},
                            { code: 'ROW([reference]) / COLUMN([reference])', desc: { en: 'Returns the row/column number of a reference.', az: 'İstinadın sətir/sütun nömrəsini qaytarır.'}},
                            { code: 'TRANSPOSE(array)', desc: { en: 'Converts a vertical range to a horizontal range, or vice versa.', az: 'Şaquli diapazonu üfüqi diapazona çevirir və ya əksinə.'}}
                        ]
                    }
                ]
            },
            'l9': {
                title: { en: 'Lesson 9: Financial Functions', az: 'Dərs 9: Maliyyə Funksiyaları' },
                topics: [
                    { title: {en: 'Core Financial Functions', az: 'Əsas Maliyyə Funksiyaları'}, desc: { en: 'Essential functions for financial analysis, investment, and loan calculations.', az: 'Maliyyə təhlili, investisiya və kredit hesablamaları üçün vacib funksiyalar.' },
                        syntax: [
                            { code: 'PV(rate, nper, pmt, [fv], [type])', desc: { en: 'Calculates the Present Value of an investment.', az: 'İnvestisiyanın Cari Dəyərini hesablayır.'}},
                            { code: 'FV(rate, nper, pmt, [pv], [type])', desc: { en: 'Calculates the Future Value of an investment.', az: 'İnvestisiyanın Gələcək Dəyərini hesablayır.'}},
                            { code: 'PMT(rate, nper, pv, [fv], [type])', desc: { en: 'Calculates the periodic Payment for a loan.', az: 'Kredit üçün dövri Ödənişi hesablayır.'}},
                            { code: 'NPER(rate, pmt, pv, [fv], [type])', desc: { en: 'Calculates the Number of Periods for an investment or loan.', az: 'İnvestisiya və ya kredit üçün Dövrlərin Sayını hesablayır.'}},
                            { code: 'RATE(nper, pmt, pv, [fv], [type])', desc: { en: 'Calculates the interest RATE per period of an annuity.', az: 'Annuitetin hər dövrü üçün faiz DƏRƏCƏSİNİ hesablayır.'}},
                            { code: 'NPV(rate, value1, [value2], ...)', desc: { en: 'Calculates the Net Present Value of an investment.', az: 'İnvestisiyanın Xalis Cari Dəyərini hesablayır.'}},
                            { code: 'IRR(values, [guess])', desc: { en: 'Calculates the Internal Rate of Return.', az: 'Daxili Gəlir Normasını hesablayır.'}},
                            { code: 'EFFECT(nominal_rate, npery)', desc: { en: 'Calculates the Effective annual interest rate.', az: 'Effektiv illik faiz dərəcəsini hesablayır.'}}
                        ]
                    }
                ]
            },
            'l10': {
                title: { en: 'Lesson 10: Data Manipulation & Cleaning Tools', az: 'Dərs 10: Data Manipulyasiyası və Təmizlənməsi Alətləri' },
                topics: [
                    { title: { en: 'Sorting & Filtering', az: 'Çeşidləmə və Filtrləmə' }, desc: { en: 'Organize data by sorting on one or multiple columns. Use filters to display only the rows that meet certain criteria.', az: 'Məlumatları bir və ya bir neçə sütuna görə çeşidləyərək təşkil edin. Yalnız müəyyən şərtlərə cavab verən sətirləri göstərmək üçün filtrlərdən istifadə edin.' }},
                    { title: { en: 'Text to Columns', az: 'Mətni Sütunlara Bölmək' }, desc: { en: 'Split the contents of one cell into separate columns based on a delimiter (like a comma or space) or fixed width.', az: 'Bir xananın məzmununu ayırıcıya (məsələn, vergül) və ya sabit enə görə ayrı-ayrı sütunlara bölün.' }},
                    { title: { en: 'Data Validation', az: 'Data Yoxlanması' }, desc: { en: 'Control what users can enter into a cell by creating dropdown lists, restricting entries to numbers, or setting date ranges.', az: 'Açılan siyahılar yaradaraq və ya girişləri məhdudlaşdıraraq istifadəçilərin bir xanaya nə daxil edə biləcəyinə nəzarət edin.' }},
                    { title: { en: 'Subtotal', az: 'Aralıq Cəm (Subtotal)' }, desc: { en: 'Insert summary rows (for SUM, COUNT, etc.) into a sorted list. It works dynamically with filters.', az: 'Çeşidlənmiş bir siyahıya yekun sətirlər (SUM, COUNT üçün) daxil edin. Bu funksiya filtrlərlə dinamik olaraq işləyir.' }},
                    { title: { en: 'Remove Duplicates & Go To Special', az: 'Dublikatları Silmək və Xüsusi Keçid' }, desc: { en: 'Permanently delete duplicate rows from a data range and use Go To Special to select specific cell types (blanks, formulas).', az: 'Təkrarlanan sətirləri həmişəlik silin və xüsusi xana növlərini (boşluqlar, formullar) seçmək üçün "Go To Special" istifadə edin.' }},
                ]
            },
            'l11': {
                title: { en: 'Lesson 11: Mastering PivotTables', az: 'Dərs 11: PivotTable-ları Mənimsəmək' },
                topics: [
                    { title: { en: 'Creating and Managing a PivotTable', az: 'PivotTable Yaratmaq və İdarə Etmək' }, desc: { en: 'Learn to create a PivotTable from a data source, refresh it, and change its source data.', az: 'Məlumat mənbəyindən PivotTable yaratmağı, onu yeniləməyi və mənbə məlumatlarını dəyişdirməyi öyrənin.' }},
                    { title: { en: 'The PivotTable Fields Pane', az: 'PivotTable Sahələr Paneli' }, desc: { en: 'Understand the four areas (Filters, Columns, Rows, Values) and how to drag and drop fields to build your report.', az: 'Dörd sahəni (Filtrlər, Sütunlar, Sətirlər, Dəyərlər) və hesabatınızı qurmaq üçün sahələri necə sürükləyəcəyinizi anlayın.' }},
                    { title: { en: 'Calculations & "Show Values As"', az: 'Hesablamalar və "Dəyərləri Belə Göstər"' }, desc: { en: 'Change the calculation from Sum to Count, Average, etc. Display values as a % of Grand Total, % of Column, or Running Total.', az: 'Hesablamanı Sum-dan Count, Average və s.-ə dəyişin. Dəyərləri Ümumi Cəmin faizi, Sütunun faizi və ya Artan Yekun kimi göstərin.' }},
                    { title: { en: 'Grouping, Sorting, and Filtering', az: 'Qruplaşdırma, Çeşidləmə və Filtrləmə' }, desc: { en: 'Group dates by month/quarter/year, sort by values, and use Slicers and Timelines for interactive filtering.', az: 'Tarixləri ay/rüb/ilə görə qruplaşdırın, dəyərlərə görə çeşidləyin və interaktiv filtrləmə üçün Slicer və Timelines-dan istifadə edin.' }},
                    { title: { en: 'Layout & Design', az: 'Düzən və Dizayn' }, desc: { en: 'Switch between Compact, Outline, and Tabular layouts. Control subtotals, grand totals, and report styling.', az: 'Kompakt, Kontur və Cədvəl düzənləri arasında keçid edin. Aralıq cəmlərə, ümumi cəmlərə və hesabat üslubuna nəzarət edin.' }}
                ]
            },
            'l12': {
                title: { en: 'Lesson 12: Data Visualization & Charts', az: 'Dərs 12: Datanın Vizuallaşdırılması və Qrafiklər' },
                topics: [
                    { title: { en: 'Choosing the Right Chart', az: 'Düzgün Qrafik Seçimi' }, desc: { en: 'Learn the principles of choosing the right chart (Column, Bar, Line, Pie) for your data and message.', az: 'Məlumatlarınız və çatdırmaq istədiyiniz mesaj üçün düzgün qrafik (Sütunlu, Zolaqlı, Xətli, Dairəvi) seçmə prinsiplərini öyrənin.' }},
                    { title: { en: 'Column, Bar & Line Charts', az: 'Sütunlu, Zolaqlı və Xətli Qrafiklər' }, desc: { en: 'Master the most common chart types for comparing categories (Column/Bar) and showing trends over time (Line).', az: 'Kateqoriyaları müqayisə etmək (Sütunlu/Zolaqlı) və zamanla trendləri göstərmək (Xətli) üçün ən çox istifadə olunan qrafik növlərini mənimsəyin.' }},
                    { title: { en: 'Combo Charts', az: 'Kombinə Edilmiş Qrafiklər' }, desc: { en: 'Combine two different chart types (e.g., column and line) on a secondary axis to show different scales of data.', az: 'Fərqli miqyaslı məlumatları göstərmək üçün iki fərqli qrafik növünü (məsələn, sütunlu və xətli) ikinci bir oxda birləşdirin.' }},
                    { title: { en: 'PivotCharts', az: 'PivotChart-lar' }, desc: { en: 'Create dynamic charts that are directly linked to a PivotTable, allowing for interactive analysis and visualization.', az: 'Birbaşa PivotTable-a bağlı olan dinamik qrafiklər yaradın, bu da interaktiv təhlilə imkan verir.' }},
                    { title: { en: 'Waterfall, Donut, and other chart types', az: 'Şəlalə, Həlqəvi və digər qrafik növləri' }, desc: { en: 'Explore other useful chart types for specific scenarios.', az: 'Xüsusi ssenarilər üçün digər faydalı qrafik növlərini araşdırın.' }}
                ]
            },
            'l13': {
                title: { en: 'Lesson 13: Worksheet & Workbook Operations', az: 'Dərs 13: İş Vərəqi və Kitabı Üzərində Əməliyyatlar' },
                topics: [
                    { title: { en: 'Protect Sheet & Workbook', az: 'İş Vərəqini və Kitabını Müdafiə Etmək' }, desc: { en: 'Prevent users from making changes to specific cells (Protect Sheet) or the workbook\'s structure (Protect Workbook).', az: 'İstifadəçilərin müəyyən xanalarda (Vərəqi Müdafiə Et) və ya iş kitabının strukturunda (Kitabı Müdafiə Et) dəyişiklik etməsinin qarşısını alın.' }},
                    { title: { en: 'Protecting Specific Ranges', az: 'Xüsusi Diapazonların Müdafiəsi' }, desc: { en: 'Lock the entire sheet but leave specific cells or ranges unlocked for user input.', az: 'Bütün vərəqi kilidləyin, ancaq istifadəçi girişi üçün müəyyən xanaları və ya diapazonları kilidsiz saxlayın.' }},
                    { title: { en: 'Page Layout & Printing Options', az: 'Səhifə Düzəni və Çap Seçimləri' }, desc: { en: 'Configure margins, orientation, and paper size. Use Print Titles to repeat header rows on every page.', az: 'Kənarları, istiqaməti və kağız ölçüsünü konfiqurasiya edin. Hər səhifədə başlıq sətirlərini təkrarlamaq üçün Çap Başlıqlarından (Print Titles) istifadə edin.' }},
                    { title: { en: 'Page Break & Layout View', az: 'Səhifə Sonu və Düzən Görünüşü' }, desc: { en: 'Visually adjust where page breaks occur and see how your worksheet will look when printed.', az: 'Səhifə sonlarının harada baş verəcəyini vizual olaraq tənzimləyin və iş vərəqinizin çapda necə görünəcəyinə baxın.' }},
                    { title: { en: 'Headers & Footers', az: 'Yuxarı və Aşağı Sərlövhələr' }, desc: { en: 'Add dynamic information like page numbers, file paths, or dates to the top or bottom of each printed page.', az: 'Hər çap edilmiş səhifənin yuxarı və ya aşağı hissəsinə səhifə nömrələri, fayl yolu və ya tarixlər kimi dinamik məlumatlar əlavə edin.' }}
                ]
            }
        }
    };
    
    // --- 2. STATE MANAGEMENT ---
    let currentLang = localStorage.getItem('syllabus_lang') || 'az';
    let currentTheme = localStorage.getItem('syllabus_theme') || 'dark';

    // --- 3. DOM ELEMENT REFERENCES ---
    const themeToggle = document.getElementById('theme-toggle');
    const langToggle = document.getElementById('lang-toggle');
    const tabNavUl = document.querySelector('.tab-nav ul');
    const tabContent = document.querySelector('.tab-content');
    
    // --- 4. RENDER FUNCTIONS ---
    function renderLesson(lessonKey) {
        const lesson = content.lessons[lessonKey];
        if (!lesson) return;

        // Clear previous content and apply fade-in animation
        tabContent.innerHTML = '';
        tabContent.style.animation = 'none';
        tabContent.offsetHeight; // Trigger reflow
        tabContent.style.animation = 'fadeIn 0.5s ease-in-out';
        
        // Create title
        const titleEl = document.createElement('h2');
        titleEl.textContent = lesson.title[currentLang];
        tabContent.appendChild(titleEl);

        // Create list of topics
        const topicsUl = document.createElement('ul');
        lesson.topics.forEach(topic => {
            const topicLi = document.createElement('li');
            
            const topicTitle = document.createElement('strong');
            topicTitle.textContent = topic.title[currentLang];
            topicLi.appendChild(topicTitle);
            
            const topicDesc = document.createElement('p');
            topicDesc.textContent = topic.desc[currentLang];
            topicLi.appendChild(topicDesc);

            // If syntax block exists for the topic, create it
            if (topic.syntax && topic.syntax.length > 0) {
                topic.syntax.forEach(syn => {
                    const syntaxDiv = document.createElement('div');
                    syntaxDiv.className = 'function-syntax';
                    
                    const codeEl = document.createElement('code');
                    codeEl.textContent = syn.code;
                    syntaxDiv.appendChild(codeEl);

                    if(syn.desc) {
                        const syntaxDescP = document.createElement('p');
                        syntaxDescP.textContent = syn.desc[currentLang];
                        syntaxDiv.appendChild(syntaxDescP);
                    }
                    
                    topicLi.appendChild(syntaxDiv);
                });
            }

            topicsUl.appendChild(topicLi);
        });
        tabContent.appendChild(topicsUl);
    }
    
    function renderTabs() {
        tabNavUl.innerHTML = '';
        Object.keys(content.lessons).forEach((key, index) => {
            const tabLi = document.createElement('li');
            tabLi.textContent = key.toUpperCase();
            tabLi.dataset.lessonKey = key;
            if (index === 0) {
                tabLi.classList.add('active');
            }
            tabNavUl.appendChild(tabLi);
        });
        addTabEventListeners();
    }
    
    function renderStaticText() {
        document.querySelector('[data-key="main_title"]').textContent = content.header.main_title[currentLang];
        document.querySelector('[data-key="prepared_by"]').textContent = content.header.prepared_by[currentLang];
    }
    
    function refreshContent() {
        renderStaticText();
        const activeTab = tabNavUl.querySelector('.active');
        if (activeTab) {
            renderLesson(activeTab.dataset.lessonKey);
        }
    }

    // --- 5. UI CONTROL & EVENT LISTENERS ---
    function applyTheme() {
        document.body.classList.toggle('dark-mode', currentTheme === 'dark');
        themeToggle.checked = (currentTheme === 'dark');
        localStorage.setItem('syllabus_theme', currentTheme);
    }

    function applyLanguage() {
        langToggle.checked = (currentLang === 'az');
        localStorage.setItem('syllabus_lang', currentLang);
        refreshContent();
    }
    
    function addTabEventListeners() {
        tabNavUl.querySelectorAll('li').forEach(tab => {
            tab.addEventListener('click', (e) => {
                tabNavUl.querySelector('.active').classList.remove('active');
                e.currentTarget.classList.add('active');
                renderLesson(e.currentTarget.dataset.lessonKey);
            });
        });
    }

    themeToggle.addEventListener('change', () => {
        currentTheme = themeToggle.checked ? 'dark' : 'light';
        applyTheme();
    });

    langToggle.addEventListener('change', () => {
        currentLang = langToggle.checked ? 'az' : 'en';
        applyLanguage();
    });

    // --- 6. INITIALIZATION ---
    function init() {
        applyTheme();
        renderTabs();
        applyLanguage(); // This also triggers the first render of content
    }
    
    init();
});
