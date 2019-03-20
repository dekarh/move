from string import digits
from collections import OrderedDict
import requests, json

from lib import read_config, l, s, fine_phone, format_phone

MANIPULATE_LABELS = ["-------------------------"
                    , "ФИО из поля"
                    , "-------------------------"
                    , "ФИО при рождении из поля"
                    , "-------------------------"
                    , "Регистрация -> Регион"
                    , "Регистрация -> Район"
                    , "Регистрация -> Город"
                    , "Регистрация -> Населенный_пункт"
                    , "Регистрация -> Улица"
                    , "-------------------------"
                    , "Проживание -> Регион"
                    , "Проживание -> Район"
                    , "Проживание -> Город"
                    , "Проживание -> Населенный_пункт"
                    , "Проживание -> Улица"
                    , "-------------------------"
                    , "Адрес регистрации из_поля"
                    , "-------------------------"
                    , "Адрес проживания из поля"
                    , "-------------------------"
                    , "Регион регистрации из номера"
                    , "-------------------------"
                    , "Регион проживания из номера"
                    , "-------------------------"
                    , "Cерия и Номер паспорта из поля"
                    , "-------------------------"
                    , "Генератор некорректных СНИЛС"
                    #, "Пол_получить_из_ФИО"
                    #, "Пол_подставить_свои_значения"
                     ]

SNILS_LABEL = ["СНИЛС"]
FIO_LABELS = ["ФИО.Фамилия", "ФИО.Имя", "ФИО.Отчество"]
FIO_BIRTH_LABELS = ["ФИО_при_рождении.Фамилия", "ФИО_при_рождении.Имя", "ФИО_при_рождении.Отчество"]
FIO_SNILS_LABELS = ["Фамилия_по_СНИЛС", "Имя_по_СНИЛС", "Отчество_по_СНИЛС"]
GENDER_LABEL = ["Пол"]
DATE_BIRTH_LABEL = ["Дата_рождения"]
PLACE_BIRTH_LABELS = ["Место_рождения.Страна", "Место_рождения.Область", "Место_рождения.Район",
                      "Место_рождения.Город"]
PASSPORT_DATA_LABELS = ["Данные_паспорта.Серия", "Данные_паспорта.Номер", "Данные_паспорта.Дата_выдачи",
                        "Данные_паспорта.Кем_выдан", "Данные_паспорта.Код_подразделения"]
ADRESS_REG_LABELS = ["Адрес_регистрации.Индекс",
                     "Адрес_регистрации.Регион", "Адрес_регистрации.Тип_региона",
                     "Адрес_регистрации.Район", "Адрес_регистрации.Тип_района",
                     "Адрес_регистрации.Город", "Адрес_регистрации.Тип_города",
                     "Адрес_регистрации.Населенный_пункт", "Адрес_регистрации.Тип_населенного_пункта",
                     "Адрес_регистрации.Улица", "Адрес_регистрации.Тип_улицы",
                     "Адрес_регистрации.Дом",
                     "Адрес_регистрации.Корпус",
                     "Адрес_регистрации.Квартира"]

ADRESS_LIVE_LABELS = ["Адрес_проживания.Индекс",
                      "Адрес_проживания.Регион", "Адрес_проживания.Тип_региона",
                      "Адрес_проживания.Район", "Адрес_проживания.Тип_района",
                      "Адрес_проживания.Город", "Адрес_проживания.Тип_города",
                      "Адрес_проживания.Населенный_пункт", "Адрес_проживания.Тип_населенного_пункта",
                      "Адрес_проживания.Улица", "Адрес_проживания.Тип_улицы",
                      "Адрес_проживания.Дом",
                      "Адрес_проживания.Корпус",
                      "Адрес_проживания.Квартира"]

PHONES_LABELS = ["Телефон.Мобильный", "Телефон.Родственников", "Телефон.Домашний"]

TECH_LABELS = ["Агент_Ид", "Подписант_Ид", "Пред_Страховщик_Ид"]

FIELDS_IN_RESULT_TABLE_FULL = [SNILS_LABEL, FIO_LABELS, FIO_BIRTH_LABELS, FIO_SNILS_LABELS, GENDER_LABEL,
                                   DATE_BIRTH_LABEL, PLACE_BIRTH_LABELS, PASSPORT_DATA_LABELS, ADRESS_REG_LABELS,
                                   ADRESS_LIVE_LABELS, PHONES_LABELS, TECH_LABELS, MANIPULATE_LABELS]
FIELDS_IN_RESULT_TABLE_SHORT = [SNILS_LABEL, FIO_LABELS, FIO_BIRTH_LABELS, FIO_SNILS_LABELS, GENDER_LABEL,
                                   DATE_BIRTH_LABEL, PLACE_BIRTH_LABELS, PASSPORT_DATA_LABELS, ADRESS_REG_LABELS,
                                   ADRESS_LIVE_LABELS, PHONES_LABELS, TECH_LABELS]

HEAD_RESULT_EXCEL_FILE = ['СНИЛС',
                          'Фамилия', 'Имя', 'Отчество',
                          'Фамилия_при_рождении', 'Имя_при_рождении', 'Отчество_при_рождении',
                          'Фамилия_по_СНИЛС', 'Имя_по_СНИЛС', 'Отчество_по_СНИЛС',
                          'Пол(0_мужской,1_женский)',
                          'Дата_рождения',
                          'Страна_рождения', 'Область_рождения', 'Район_рождения', 'Город_рождения',
                          'Паспорт_серия', 'Паспорт_номер', 'Паспорт_дата', 'Паспорт_Кем выдан',
                          'Паспорт_Код подразделения',

                          'Адрес_регистрации_Индекс',
                          'Адрес_регистрации_Регион', 'Адрес_регистрации_Тип_региона',
                          'Адрес_регистрации_Район', 'Адрес_регистрации_Тип_района',
                          'Адрес_регистрации_Город', 'Адрес_регистрации_Тип_города',
                          'Адрес_регистрации_Населенный_пункт', 'Адрес_регистрации_Тип_населенного_пункта',
                          'Адрес_регистрации_Улица',
                          'Адрес_регистрации_Тип_улицы',
                          'Адрес_регистрации_Дом',
                          'Адрес_регистрации_Корпус',
                          'Адрес_регистрации_Квартира',

                          'Адрес_проживания_Индекс',
                          'Адрес_проживания_Регион', 'Адрес_проживания_Тип_региона',
                          'Адрес_проживания_Район', 'Адрес_проживания_Тип_района',
                          'Адрес_проживания_Город', 'Адрес_проживания_Тип_города',
                          'Адрес_проживания_Населенный_пункт', 'Адрес_проживания_Тип_населенного_пункта',
                          'Адрес_проживания_Улица', 'Адрес_проживания_Тип_улицы',
                          'Адрес_проживания_Дом',
                          'Адрес_проживания_Корпус',
                          'Адрес_проживания_Квартира',

                          'Мобильный_телефон', 'Телефон_родственников', 'Телефон_домашний',
                          'Агент_Ид', 'Подписант_Ид', 'Пред_Страховщик_Ид'
                          ]

LEN_SNILS = 11
LEN_PASSPORT_NOMER = 6
LEN_INDEX_NOMER = 6
LEN_PASSPORT_COD = 6
EAST_GENDER = ['кызы', 'оглы']

########################################################################################################################
# НУЛЕВЫЕ ЗНАЧЕНИЯ
NULL_VALUE = '\\N'  # НУЛЕВОЕ ЗНАЧЕНИЕ В ФАЙЛЕ
NEW_NULL_VALUE = ''  # НОВОЕ НУЛЕВОЕ ЗНАЧЕНИЕ

NEW_NULL_VALUE_FOR_DATE = '11.11.1111'
NEW_NULL_VALUE_FOR_SERIYA_PASSPORTA = '1111'
NEW_NULL_VALUE_FOR_NOMER_PASSPORTA = '111111'
NEW_NULL_VALUE_FOR_COD_PASSPORTA = '111-111'
NEW_NULL_VALUE_FOR_GENDER = '0'
NEW_NULL_VALUE_FOR_INDEX = '111111'
NEW_NULL_VALUE_FOR_ALL_TEXT = 'заполнить'
NEW_NULL_VALUE_FOR_HOME = 'заполнить'
########################################################################################################################
# ЗНАЧЕНИЕ ПРИ ОШИБКЕ
ERROR_VALUE = 'ERROR'
########################################################################################################################
# СОКРАЩЕНИЯ ТИПОВ В АДРЕСЕ
SPLIT_FIELDS = ['.', ',', ' ', ';', '№']
SPLIT_FIELD = SPLIT_FIELDS[1]
#SPLIT_FIELD = '_x0003_'  # Разделитель для адреса в одной строке (бывает '_x0003_')

# Для варианта когда все поля по раздельности и перепутаны. Используется в FullAdress.get_values. Справочно:
# ['Индекс', 'Регион', 'Тип_региона', 'Район', 'Тип_района', 'Город', 'Тип_города',
#  'Населенный_пункт', 'Тип_населенного_пункта', 'Улица', 'Тип_улицы', 'Дом', 'Корпус', 'Квартира']
ORDER_FIELD = [13, 0, 8, 1, 9, 2, 10, 3, 11, 4, 12, 5, 6, 7]

REG_TYPES = ['обл', 'о', 'область', 'респ', 'республика', 'край', 'кр', 'ар', 'ао', 'авт окр', 'автономный округ',
             'авт обл', 'автономная область', 'город федерального значения', 'гфз']

DISTRICT_TYPES = ['р-н', 'р', 'район']

CITY_TYPES = ['г', 'гор', 'город']

NP_TYPES = ['пгт', 'поселок городского типа',  'посёлок городского типа', 'пос', 'поселение', 'поселок', 'посёлок',
            'п', 'рп', 'рабочий посёлок', 'рабочий поселок', 'кп', 'курортный посёлок', 'курортный поселок', 'к', 'пс',
            'сс', 'смн', 'вл', 'влад', 'владение', 'дп', 'дачный поселок', 'дачный посёлок', 'садовое товарищество',
            'садоводческое некоммерческое товарищество', 'садоводческое товарищество', 'снт', 'нп', 'пст', 'ж/д_ст',
            'ж/д ст', 'железнодорожная станция', 'с', 'село', 'м', 'д', 'дер', 'деревня', 'сл', 'ст', 'ст-ца',
            'станица', 'х', 'хут', 'хутор', 'рзд', 'у', 'урочище', 'клх', 'колхоз', 'свх', 'совхоз', 'зим', 'зимовье',
            'микрорайон', 'мкр']

STREET_TYPES = ['аллея', 'а', 'бульвар', 'б-р', 'бул', 'в/ч', 'военная часть', 'военный городок', 'городок', 'гск',
                'гаражно-строительный кооператив', 'гк', 'гаражный кооператив', 'кв-л', 'квартал', 'линия', 'лин',
                'наб', 'набережная', 'переулок', 'пер', 'переезд', 'пл', 'площадь', 'пр-кт', 'проспект', 'пр',
                'проезд', 'тер', 'терр', 'территория', 'туп', 'тупик', 'ул', 'улица', 'ш', 'шоссе']

HOUSE_CUT_NAME = ['дом', 'д']
CORPUS_CUT_NAME = ['корп', 'корпус', 'стр', 'строение']
APARTMENT_CUT_NAME = ['кв', 'квартира', 'оф', 'офис', 'ап', 'аппартаменты']

ADRESS_TYPES = {
'обл': 1, 'о': 1, 'область': 1, 'респ': 1, 'республика': 1, 'край': 1, 'кр': 1, 'ар': 1, 'ао': 1, 'авт окр': 1, 'автономный округ': 1, 'авт обл': 1, 'автономная область': 1, 'город федерального значения': 1, 'гфз': 1,
'р-н': 3, 'р': 3, 'район': 3,
'г': 5, 'гор': 5, 'город': 5,
'пгт': 7, 'поселок городского типа': 7,  'посёлок городского типа': 7, 'пос': 7, 'поселение': 7, 'поселок': 7, 'посёлок': 7, 'п': 7, 'рп': 7, 'рабочий посёлок': 7, 'рабочий поселок': 7, 'кп': 7, 'курортный посёлок': 7, 'курортный поселок': 7, 'к': 7, 'пс': 7, 'сс': 7, 'смн': 7, 'вл': 7, 'влад': 7, 'владение': 7, 'дп': 7, 'дачный поселок': 7, 'дачный посёлок': 7, 'садовое товарищество': 7, 'садоводческое некоммерческое товарищество': 7, 'садоводческое товарищество': 7, 'снт': 7, 'нп': 7, 'пст': 7, 'ж/д_ст': 7, 'ж/д ст': 7, 'железнодорожная станция': 7, 'с': 7, 'село': 7, 'м': 7, 'дер': 7, 'деревня': 7, 'сл': 7, 'ст': 7, 'ст-ца': 7, 'станица': 7, 'х': 7, 'хут': 7, 'хутор': 7, 'рзд': 7, 'у': 7, 'урочище': 7, 'клх': 7, 'колхоз': 7, 'свх': 7, 'совхоз': 7, 'зим': 7, 'зимовье': 7, 'микрорайон': 7, 'мкр' : 7,
'аллея': 9, 'а': 9, 'алл': 9, 'бульвар': 9, 'б-р': 9, 'бул': 9, 'блв': 9, 'в/ч': 9, 'военная часть': 9, 'военный городок': 9, 'городок': 9, 'гск': 9, 'гаражно-строительный кооператив': 9, 'гк': 9, 'гаражный кооператив': 9, 'кв-л': 9, 'квартал': 9, 'линия': 9, 'лин': 9, 'наб': 9, 'набережная': 9, 'переулок': 9, 'пер': 9, 'переезд': 9, 'пл': 9, 'площадь': 9, 'пр-кт': 9, 'проспект': 9, 'пр': 9, 'проезд': 9, 'тер': 9, 'терр': 9, 'территория': 9, 'туп': 9, 'тупик': 9, 'ул': 9, 'улица': 9, 'ш': 9, 'шоссе': 9,
'дом': 11,  'д': 11,
'корп': 12,  'корпус': 12,  'стр': 12,  'строение': 12,
'кв': 13,  'квартира': 13,  'оф': 13,  'офис': 13,  'ап': 13,  'аппартаменты': 13
}

ALL_CUT_NAMES = [
    'а обл','а окр','АО','Аобл','г','г','г ф з','гфз','край','обл','обл','округ','Респ','респ','АО','АО','вн тер г',
    'г о', 'го','м р-н','п','пос','р-н','тер','у','у','волость','г','г','дп','кп','массив','п','п/о','пгт','пгт','рп',
    'с/а','с/мо','с/о','с/п','с/с','тер','р-н','тер','аал','автодорога','арбан','аул','волость','высел','г','г-к','гп',
    'д','дп','ж/д б-ка','ж/д пл-ка','ж/д пл-ма','ж/д_будка','ж/д_казарм','ж/д_оп','ж/д_платф','ж/д_пост','ж/д_рзд',
    'ж/д_ст','жилзона','жилрайон','заимка','казарма','кв-л','кордон','кп','лпх','м','массив','мкр','нп','остров','п',
    'п/о','п/р','п/ст','пгт','пгт','погост','починок','промзона','рзд','рп','с','сл','снт','ст','ст-ца','тер','у','х',
    'а/я','ал','аллея','балка','берег','б-р','бугор','вал','взв','въезд','г-к','гск','д','днп','дор','ж/д_будка',
    'ж/д_казарм','ж/д_оп','ж/д_платф','ж/д_пост','ж/д_рзд','ж/д_ст','жт','заезд','ззд','зона','казарма','кв-л','км',
    'кольцо','коса','к-цо','линия','лн','м','маяк','мгстр','местность','мкр','мост','н/п','наб','наб','нп','остров',
    'п','п/о','п/р','п/ст','парк','пер','пер','пер-д','переезд','пл','платф','пл-ка','полустанок','пр-д','пр-к','пр-ка'
    ,'пр-кт','пр-лок','проезд','промзона','просек','просека','проселок','проул','проулок','рзд','рзд','ряд','ряды','с',
    'с/т','сад','сзд','с-к','сквер','сл','снт','спуск','с-р','ст','стр','тер','тер ДНТ','тер СНТ','тракт','туп','ул',
    'ул','уч-к','ф/х','ферма','х','ш','ш','влд','д','двлд','ДОМ','зд','к','кот','ОНС','пав','соор','стр','шахта','г-ж',
    'кв','ком','офис','п-б','подв','помещ','раб уч','скл','торг зал','цех','вн р-н','г п','с п','с/с','а/я','аал','ал',
    'арбан','аул','б-г','б-р','вал','взд','г-к','гск','д','днп','дор','ж/д б-ка','ж/д к-ма','ж/д пл-ма','ж/д рзд',
    'ж/д ст','ж/р','ззд','зона','кв-л','км','коса','к-цо','лн','местность','месторожд','м-ко','мкр','мкр','н/п','наб',
    'ост-в','п','п/р','парк','пер','пер-д','п-к','пл','платф','пл-ка','порт','пр-д','пр-к','пр-ка','пр-кт','пр-лок',
    'промзона','проул','рзд','р-н','с','сад','с-к','сквер','сл','снт','с-р','ст','стр','тер','тер','тер ГСК','тер ДНО',
    'тер ДНП','тер ДНТ','тер ДПК','тер ОНО','тер ОНП','тер ОНТ','тер ОПК','тер СНО','тер СНП','тер СНТ','тер СПК',
    'тер ТСН','тер СОСН','тер ф х','тракт','туп','ул','ус','ф/х','х','ш','ю','з/у','гск','днп','местность','мкр','н/п',
    'промзона','сад','снт','тер','ф/х','а/я','аал','аллея','арбан','аул','берег','б-р','вал','въезд','высел','г-к',
    'гск','д','дор','ж/д_будка','ж/д_казарм','ж/д_оп','ж/д_платф','ж/д_пост','ж/д_рзд','ж/д_ст','жт','заезд','зона',
    'казарма','кв-л','км','кольцо','коса','линия','м','мкр','мост','наб','нп','остров','п','п/о','п/р','п/ст','парк',
    'пер','переезд','пл','платф','пл-ка','починок','пр-кт','проезд','просек','просека','проселок','проулок','рзд',
    'ряды','с','сад','сквер','сл','снт','спуск','ст','стр','тер','тракт','туп','ул','уч-к','ферма','х','ш'
]

q1 = """
'обл': 'REG_TYPES', 'о': 'REG_TYPES', 'область': 'REG_TYPES', 'респ': 'REG_TYPES', 'республика': 'REG_TYPES', 'край': 'REG_TYPES', 'кр': 'REG_TYPES', 'ар': 'REG_TYPES', 'ао': 'REG_TYPES', 'авт окр': 'REG_TYPES', 'автономный округ': 'REG_TYPES', 'авт обл': 'REG_TYPES', 'автономная область': 'REG_TYPES', 'город федерального значения': 'REG_TYPES', 'гфз': 'REG_TYPES',
'р-н': 'DISTRICT_TYPES', 'р': 'DISTRICT_TYPES', 'район': 'DISTRICT_TYPES',
'г': 'CITY_TYPES', 'гор': 'CITY_TYPES', 'город': 'CITY_TYPES',
'пгт': 'NP_TYPES', 'поселок городского типа': 'NP_TYPES',  'посёлок городского типа': 'NP_TYPES', 'пос': 'NP_TYPES', 'поселение': 'NP_TYPES', 'поселок': 'NP_TYPES', 'посёлок': 'NP_TYPES', 'п': 'NP_TYPES', 'рп': 'NP_TYPES', 'рабочий посёлок': 'NP_TYPES', 'рабочий поселок': 'NP_TYPES', 'кп': 'NP_TYPES', 'курортный посёлок': 'NP_TYPES', 'курортный поселок': 'NP_TYPES', 'к': 'NP_TYPES', 'пс': 'NP_TYPES', 'сс': 'NP_TYPES', 'смн': 'NP_TYPES', 'вл': 'NP_TYPES', 'влад': 'NP_TYPES', 'владение': 'NP_TYPES', 'дп': 'NP_TYPES', 'дачный поселок': 'NP_TYPES', 'дачный посёлок': 'NP_TYPES', 'садовое товарищество': 'NP_TYPES', 'садоводческое некоммерческое товарищество': 'NP_TYPES', 'садоводческое товарищество': 'NP_TYPES', 'снт': 'NP_TYPES', 'нп': 'NP_TYPES', 'пст': 'NP_TYPES', 'ж/д_ст': 'NP_TYPES', 'ж/д ст': 'NP_TYPES', 'железнодорожная станция': 'NP_TYPES', 'с': 'NP_TYPES', 'село': 'NP_TYPES', 'м': 'NP_TYPES', 'дер': 'NP_TYPES', 'деревня': 'NP_TYPES', 'сл': 'NP_TYPES', 'ст': 'NP_TYPES', 'ст-ца': 'NP_TYPES', 'станица': 'NP_TYPES', 'х': 'NP_TYPES', 'хут': 'NP_TYPES', 'хутор': 'NP_TYPES', 'рзд': 'NP_TYPES', 'у': 'NP_TYPES', 'урочище': 'NP_TYPES', 'клх': 'NP_TYPES', 'колхоз': 'NP_TYPES', 'свх': 'NP_TYPES', 'совхоз': 'NP_TYPES', 'зим': 'NP_TYPES', 'зимовье': 'NP_TYPES', 'микрорайон': 'NP_TYPES', 'мкр' : 'NP_TYPES',
'аллея': 'STREET_TYPES', 'а': 'STREET_TYPES', 'алл': 'STREET_TYPES', 'бульвар': 'STREET_TYPES', 'б-р': 'STREET_TYPES', 'бул': 'STREET_TYPES', 'блв': 'STREET_TYPES', 'в/ч': 'STREET_TYPES', 'военная часть': 'STREET_TYPES', 'военный городок': 'STREET_TYPES', 'городок': 'STREET_TYPES', 'гск': 'STREET_TYPES', 'гаражно-строительный кооператив': 'STREET_TYPES', 'гк': 'STREET_TYPES', 'гаражный кооператив': 'STREET_TYPES', 'кв-л': 'STREET_TYPES', 'квартал': 'STREET_TYPES', 'линия': 'STREET_TYPES', 'лин': 'STREET_TYPES', 'наб': 'STREET_TYPES', 'набережная': 'STREET_TYPES', 'переулок': 'STREET_TYPES', 'пер': 'STREET_TYPES', 'переезд': 'STREET_TYPES', 'пл': 'STREET_TYPES', 'площадь': 'STREET_TYPES', 'пр-кт': 'STREET_TYPES', 'проспект': 'STREET_TYPES', 'пр': 'STREET_TYPES', 'проезд': 'STREET_TYPES', 'тер': 'STREET_TYPES', 'терр': 'STREET_TYPES', 'территория': 'STREET_TYPES', 'туп': 'STREET_TYPES', 'тупик': 'STREET_TYPES', 'ул': 'STREET_TYPES', 'улица': 'STREET_TYPES', 'ш': 'STREET_TYPES', 'шоссе': 'STREET_TYPES',
'дом': 'HOUSE_CUT_NAME',  'д': 'HOUSE_CUT_NAME',
'корп': 'CORPUS_CUT_NAME',  'корпус': 'CORPUS_CUT_NAME',  'стр': 'CORPUS_CUT_NAME',  'строение': 'CORPUS_CUT_NAME',
'кв': 'APARTMENT_CUT_NAME',  'квартира': 'APARTMENT_CUT_NAME',  'оф': 'APARTMENT_CUT_NAME',  'офис': 'APARTMENT_CUT_NAME',  'ап': 'APARTMENT_CUT_NAME',  'аппартаменты': 'APARTMENT_CUT_NAME'
"""
########################################################################################################################
# ЗНАЧЕНИЕ В ПОЛЕ "ПОЛ" ИЗМЕНЯЕМ В ПРОЦЕССЕ
female_gender_value = 'Ж'
male_gender_value = 'М'
gender_length = 1
########################################################################################################################
# ЗАПОЛНЕНИЕ Агент_Ид, Подписант_Ид, Пред_Страховщик_Ид
#AGENT_ID = '10061'
#AGENT_ID = '9954'
#AGENT_ID = '9986'
#PODPISANT_ID = '208'
PREDSTRAH_ID = '1'
########################################################################################################################
# ИМЕНА ДЛЯ КЛЮЧЕЙ СЛОВАРЕЙ И ДЛЯ ПОРЯДКА ВЫВОД СЛОВАРЯ

FULL_ADRESS_LABELS = ['Индекс', 'Регион', 'Тип_региона', 'Район', 'Тип_района', 'Город', 'Тип_города',
                      'Населенный_пункт', 'Тип_населенного_пункта', 'Улица', 'Тип_улицы', 'Дом', 'Корпус', 'Квартира']

PASSPORT_LABELS = ['Серия', 'Номер', 'Дата_выдачи', 'Кем_выдан', 'Код_подразделения']

FIO_KEY_LABELS = ['Фамилия', 'Имя', 'Отчество']

BIRTH_PLACE_LABELS = ['Страна', 'Область', 'Район', 'Город']
########################################################################################################################
REGIONS = [
    "", "Адыгея респ.", "Башкортостан респ.", "Бурятия респ.", "Алтай респ.", "Дагестан респ.", "Ингушетия респ.",
    "Кабардино-Балкарская респ.", "Калмыкия респ.", "Карачаево-Черкесская респ.", "Карелия респ.", "Коми респ.",
    "Марий_Эл респ.", "Мордовия респ.", "Саха/Якутия/ респ.", "Северная_Осетия-Алания респ.", "Татарстан респ.",
    "Тыва респ.", "Удмуртская респ.", "Хакасия респ.", "Чеченская респ.", "Чувашская респ.", "Алтайский край",
    "Краснодарский край", "Красноярский край", "Приморский край", "Ставропольский край", "Хабаровский край",
    "Амурская обл.", "Архангельская обл.", "Астраханская обл.", "Белгородская обл.", "Брянская обл.",
    "Владимирская обл.", "Волгоградская обл.", "Вологодская обл.", "Воронежская обл.", "Ивановская обл.",
    "Иркутская обл.", "Калининградская обл.", "Калужская обл.", "Камчатский край", "Кемеровская обл.", "Кировская обл.",
    "Костромская обл.", "Курганская обл.", "Курская обл.", "Ленинградская обл.", "Липецкая обл.", "Магаданская обл.",
    "Московская обл.", "Мурманская обл.", "Нижегородская обл.", "Новгородская обл.", "Новосибирская обл.",
    "Омская обл.", "Оренбургская обл.", "Орловская обл.", "Пензенская обл.", "Пермский край", "Псковская обл.",
    "Ростовская обл.", "Рязанская обл.", "Самарская обл.", "Саратовская обл.", "Сахалинская обл.", "Свердловская обл.",
    "Смоленская обл.", "Тамбовская обл.", "Тверская обл.", "Томская обл.", "Тульская обл.", "Тюменская обл.",
    "Ульяновская обл.", "Челябинская обл.", "Забайкальский край", "Ярославская обл.", "Москва г.", "Санкт-Петербург г.",
    "Еврейская авт.обл.", "Агинский_Бурятский авт.округ", "Коми-Пермяцкий авт.округ", "Корякский авт.округ",
    "Ненецкий авт.округ", "Таймырский_(Долгано-Ненецкий) авт.округ", "Усть-Ордынский_Бурятский авт.округ",
    "Ханты-Мансийский/Югра/ авт.округ", "Чукотский авт.округ", "Эвенкийский авт.округ", "Ямало-Ненецкий авт.округ",
    "","Крым респ.", "Севастополь г.","","","","","","","Байконур г." ]
########################################################################################################################
# True - заменять СНИЛС на некорректный, неиспользованный в Сатурн

GENERATE_SNILS = False

########################################################################################################################

IN_IDS = ['ID','ИД_КЛИЕНТА','CLIENT_ID']
IN_SNILS = ['СНИЛС', 'СТРАХОВОЙ_НОМЕР', 'СТРАХОВОЙ НОМЕР', 'NUMBER','СТРАХОВОЙНОМЕР']
IN_NAMES = ['ID', 'СНИЛС', 'СТРАХОВОЙ_НОМЕР', 'СТРАХОВОЙ НОМЕР', 'NUMBER', 'ФАМИЛИЯ', 'ИМЯ', 'ОТЧЕСТВО', 'ФИО']

DIR4MOVE = '/home/da3/Move/'
DIR4IMPORT = '/home/da3/CheckLoad/'
DIR4CFGIMPORT = '/home/da3/CheckLoad/cfg/'
DIR4PCHECK = '/home/da3/PasportChecks/'


class BaseClass:
    def __setattr__(self, name, value):
        if isinstance(value, (int, str)):
            self.__dict__[name] = str(value).strip()
        else:
            self.__dict__[name] = value


class FullAdress(BaseClass):
    def __init__(self, field='', tip='по типам субъектов'):
    #def __init__(self, field='', tip='стандартный'):
        self.field = str(field)
        self.field_home = ''
        self.homes = []
        self.index = []
        self.tip = tip
        self.full_adress = []
        self.FULL_ADRESS_DICT = {}
        for label in FULL_ADRESS_LABELS:
            self.FULL_ADRESS_DICT[label] = ''
        self.iter_types = [DISTRICT_TYPES, CITY_TYPES, NP_TYPES, STREET_TYPES, HOUSE_CUT_NAME, CORPUS_CUT_NAME, APARTMENT_CUT_NAME]

    def normalize_adress(self):
        if len(self.field) != 0 and self.field != NULL_VALUE:
            self.field = self.field.lower()
            values = self.field.split(SPLIT_FIELD) # разделили на массив разделителем SPLIT_FIELD
            for i, word in enumerate(values):
                n = []
                word = word.strip()
                if i == 0:
                    n = [char for char in word if char in digits]
                    if len(n) != 6:
                        return NEW_NULL_VALUE_FOR_INDEX
                    self.FULL_ADRESS_DICT[FULL_ADRESS_LABELS[0]] = ''.join(n)
                    continue
                elif i == 1:
                    self.FULL_ADRESS_DICT[FULL_ADRESS_LABELS[1]] = ' '.join(word.split(' ')[:-1])
                    self.FULL_ADRESS_DICT[FULL_ADRESS_LABELS[2]] = word.split(' ')[-1]
                    continue
                else:
                    for j, types in enumerate(self.iter_types):
                        if j < 4:
                            if word.split(' ')[-1] in types:
                                self.FULL_ADRESS_DICT[FULL_ADRESS_LABELS[3 + 2 * j]] = ' '.join(
                                    word.split(' ')[:-1])
                                self.FULL_ADRESS_DICT[FULL_ADRESS_LABELS[3 + 2 * j + 1]] = word.split(' ')[-1]
                        elif j >= 4:
                            for type in types:
                                if word.find(type) != (-1):
                                    word = word.replace(type, '').replace('.', '')
                                    self.FULL_ADRESS_DICT[FULL_ADRESS_LABELS[11 + j - 4]] = word
            return self.FULL_ADRESS_DICT
        else:
            if self.field == NULL_VALUE:
                return NEW_NULL_VALUE
            else:
                return NEW_NULL_VALUE
                #                return ERROR_VALUE

    def create_output_list(self):                   # Создаем список вывода
        if self.field != '':
            FULL_ADRESS_DICT = self.normalize_adress()
        for label in FULL_ADRESS_LABELS:
            self.full_adress.append(self.FULL_ADRESS_DICT[label].upper())
        return self.full_adress

    def cut_adress(self):
        # Находим все разделители полей !!!!! ПЕРЕДЕЛАЛ ПОД ДУБЛИ !!!!!!
        self.field = self.field.lower()
        breaks = {}
        breaks_len = {}
        breaks_name = {}
        for adress_type in ADRESS_TYPES:
            for left in SPLIT_FIELDS:
                for right in SPLIT_FIELDS:
                    if len(self.field.split(left + adress_type + right)) > 1:
                        try:
                            tek_pos = 0
                            for i, tek_part in enumerate(self.field.split(left + adress_type + right)):
                                if i == len(self.field.split(left + adress_type + right)) - 1:
                                    break
                                tek_pos += len(tek_part) + len(left + adress_type + right)
                                breaks[ADRESS_TYPES[adress_type]].append(tek_pos)
                        except KeyError:
                            breaks[ADRESS_TYPES[adress_type]] = []
                            tek_pos = 0
                            for i, tek_part in enumerate(self.field.split(left + adress_type + right)):
                                if i == len(self.field.split(left + adress_type + right)) - 1:
                                    break
                                tek_pos += len(tek_part) + len(left + adress_type + right)
                                breaks[ADRESS_TYPES[adress_type]].append(tek_pos)
                        try:
                            for i, tek_part in enumerate(self.field.split(left + adress_type + right)):
                                if i == len(self.field.split(left + adress_type + right)) - 1:
                                    break
                                breaks_name[ADRESS_TYPES[adress_type]].append(adress_type)
                        except KeyError:
                            breaks_name[ADRESS_TYPES[adress_type]] = []
                            for i, tek_part in enumerate(self.field.split(left + adress_type + right)):
                                if i == len(self.field.split(left + adress_type + right)) - 1:
                                    break
                                breaks_name[ADRESS_TYPES[adress_type]].append(adress_type)

        digits_count = 0
        digits_pos = 0
        index_pos = -1
        for i, char in enumerate(self.field):
            if char in digits:
                if i - digits_pos > 1 or digits_count > 6:
                    digits_count = 1
                else:
                    digits_count += 1
                digits_pos = i
                if digits_count == 6:
                    index_pos = i - 5
                    try:
                        breaks[0].append(index_pos)
                    except KeyError:
                        breaks[0] = []
                        breaks[0].append(index_pos)
                    try:
                        breaks_len[0].append(6)
                    except KeyError:
                        breaks_len[0] = []
                        breaks_len[0].append(6)
                    try:
                        breaks_name[0].append('')
                    except KeyError:
                        breaks_name[0] = []
                        breaks_name[0].append('')

        # сортируем по значениям словаря
        breaks_sorted = OrderedDict(sorted(breaks.items(), key=lambda t: t[1]))

        # Отсечь дом-корпус-квартиру
        cutted_begin = []
        cutted_end = []
        home_cut = 0
        home_pass = 0
        tek_home = ''
        last_i = 0
        for i in range(5):
            #last_break_sorted = list(breaks_sorted.keys())[0]
            for break_sorted in breaks_sorted:
                if break_sorted == 0 and len(breaks_sorted[break_sorted]) > i: # индекс
                    try:
                        if home_cut: # дом-корпус-квартира (окончание)
                            cutted_end.append(breaks_sorted[break_sorted][i])
                            tek_home += ', ' + self.field[home_pass:breaks_sorted[break_sorted][i]].strip()
                            home_cut = 0
                            self.homes.append(tek_home)
                            tek_home = ''
                            #self.homes.append(self.field[cutted_begin[len(cutted_begin) - 1]:
                            #                             cutted_end[len(cutted_end) - 1]].strip())
                        cutted_begin.append(breaks_sorted[break_sorted][i])
                        cutted_end.append(breaks_sorted[break_sorted][i] + breaks_len[break_sorted][i])
                        self.index.append(self.field[breaks_sorted[break_sorted][i]:breaks_sorted[break_sorted][i] +
                                                                                    breaks_len[break_sorted][i]])
                        last_break_sorted = break_sorted
                    except IndexError:
                        pass
                elif break_sorted == 11 and len(breaks_sorted[break_sorted]) > i: # дом-корпус-квартира (начало)
                    try:
                        cutted_begin.append(breaks_sorted[break_sorted][i] - len(breaks_name[break_sorted][i]) - 1)
                        home_cut = breaks_sorted[break_sorted][i]
                        home_pass = breaks_sorted[break_sorted][i] - len(breaks_name[break_sorted][i]) - 1
                        last_break_sorted = break_sorted
                    except IndexError:
                        pass
                elif break_sorted > 11 and len(breaks_sorted[break_sorted]) > i:
                    if home_cut:
                        try:
                            tek_home += ', ' + self.field[home_pass:breaks_sorted[break_sorted][i]].strip(
                                               )[:-len(breaks_name[break_sorted][i])].strip()
                            home_pass = breaks_sorted[break_sorted][i] - len(breaks_name[break_sorted][i]) - 1
                        except IndexError:
                            pass
                    last_break_sorted = break_sorted
                elif break_sorted < 11 and len(breaks_sorted[break_sorted]) > i: # дом-корпус-квартира (окончание)
                    if home_cut:
                        try:
                            cutted_end.append(breaks_sorted[last_break_sorted][last_i] + len(tek_part))
                            tek_home += ', ' + self.field[home_pass:breaks_sorted[last_break_sorted][last_i] + len(tek_part)].strip()
                            home_cut = 0
                            self.homes.append(tek_home)
                            tek_home = ''
                        except IndexError:
                            pass
                    last_break_sorted = break_sorted
            last_i = i
        if len(cutted_begin) != len(cutted_end):
            cutted_end.append(len(self.field))
            tek_home += ', ' + self.field[home_pass:].strip()
            home_cut = 0
            self.homes.append(tek_home)

            #self.homes.append(self.field[cutted_begin[len(cutted_begin) - 1]:cutted_end[len(cutted_end) - 1]].strip())

        # удалить индекс, дом-корпус-квартиру
        field_cut = ''
        if len(cutted_begin):
            field_cut = self.field[:cutted_begin[0]]
            for i, cutted in enumerate(cutted_begin):
                if i:
                    field_cut += self.field[cutted_end[i-1]:cutted]

        # Заменить разделители на пробелы, схлопнуть двойные пробелы в одинарные
        for cut in SPLIT_FIELDS:
            field_cut = field_cut.replace(cut, ' ')
        while field_cut.find('  ') > -1:
            field_cut = field_cut.replace('  ', ' ')

        # Удалить типы объектов
        for adress_type in ADRESS_TYPES:
            field_cut = field_cut.replace(' ' + adress_type + ' ',' ')
        for adress_type in ALL_CUT_NAMES:
            field_cut = field_cut.replace(' ' + adress_type + ' ', ' ')
        while field_cut.find('  ') > -1:
            field_cut = field_cut.replace('  ', ' ')

        # Удалить дублирующиеся слова
        words = field_cut.split()
        field_cut = ''
        for word in words:
            if word not in field_cut.split():
                field_cut += word + ' '
        self.field_home = field_cut

    def get_values(self):
        if self.tip == 'стандартный':
            # Когда адрес 414000, г. Астрахань, ул. Такая, д. Т...
            output_list = []
            for elem in self.create_output_list():
                output_list.append(elem.strip())
            return output_list
        elif self.tip == 'перемешаный':
            # Когда все поля по раздельности и перемешаны...
            output_list = []

            if len(self.field) != 0 and self.field != NULL_VALUE:
                self.field = self.field.lower()
                values = self.field.split(SPLIT_FIELD)
                for i, nn in enumerate(ORDER_FIELD):
                    if nn < len(values):
                        output_list.append(values[nn])
                return output_list
            else:
                if self.field == NULL_VALUE:
                    return NEW_NULL_VALUE
                else:
                    return NEW_NULL_VALUE
        elif self.tip == 'по типам субъектов':
            # Заменить разделители на пробелы, схлопнуть двойные пробелы в одинарные
            field_cut = self.field
            for cut in SPLIT_FIELDS:
                field_cut = field_cut.replace(cut, ' ')
            while field_cut.find('  ') > -1:
                field_cut = field_cut.replace('  ', ' ')
            self.field = field_cut
            output_list = []
            breaked_address = False
            if len(self.field) > 0 and self.field != NULL_VALUE:
                self.cut_adress()
                if len(self.homes):
                    base_home = self.homes[0]
                    all_homes = self.homes[0]
                    for home in self.homes:
                        if home != base_home:
                            all_homes += home
                            breaked_address = True
                if len(self.index):
                    base_index = self.index[0]
                    all_index = self.index[0]
                    for index in self.index:
                        if index != base_index:
                            all_index += ', ' + index
                            breaked_address = True
                else:
                    all_index = '111111'
                res = requests.get('http://127.0.0.1:23332/find/' + self.field_home + '?strong=1')
                if res.status_code == 200:
                    ajson = json.loads(bytes.decode(res.content))
                    print(all_index + ', ' + ajson[0]['text'] + all_homes)

                breaks = {}
                breaks_len = {}
                breaks_name = {}
                for adress_type in ADRESS_TYPES:
                    for left in SPLIT_FIELDS:
                        for right in SPLIT_FIELDS:
                            if len(self.field_home.split(left + adress_type + right)) > 1:
                                breaks[ADRESS_TYPES[adress_type]] = self.field_home.find(left + adress_type + right)
                                breaks_name[ADRESS_TYPES[adress_type]] = adress_type
                digits_count = 0
                digits_pos = 0
                index_pos = -1
                for i, char in enumerate(self.field_home):
                    if char in digits:
                        if i - digits_pos > 1 or digits_count > 6:
                            digits_count = 1
                        else:
                            digits_count += 1
                        digits_pos = i
                        if digits_count == 6:
                            index_pos = i - 5

                if index_pos > -1:
                    breaks[0] = index_pos + 6
                    breaks_len[0] = 6
                    breaks_name[0] = ''

                # сортируем по значениям словаря
                breaks_sorted = OrderedDict(sorted(breaks.items(), key=lambda t: t[1]))
                break_sorted_last = -1
                output_dict = {}
                for i, break_sorted in enumerate(breaks_sorted):
                    if break_sorted == 0: # индекс
                        output_dict[break_sorted] = self.field_home[breaks_sorted[break_sorted] -
                                                                  breaks_len[break_sorted]: breaks_sorted[break_sorted]]
                        break_sorted_last = break_sorted
                        continue
                    # ищем значение перед типом субъекта
                    if break_sorted_last == -1:
                        subject = self.field_home[:breaks_sorted[break_sorted]]
                        breaks_len[break_sorted] = len(breaks_name[break_sorted_last])
                    else:
                        subject = self.field_home[breaks_sorted[break_sorted_last] + len(breaks_name[break_sorted_last])
                                                  + 1: breaks_sorted[break_sorted]]
                        breaks_len[break_sorted] = len(breaks_name[break_sorted])
                    for j in range(5):  # тщательно обрезаем разделительные символы с концов строки
                        for split_field in SPLIT_FIELDS:
                            subject = subject.strip(split_field)
                    if len(subject):    # что-нибудь осталось?
                        output_dict[break_sorted] = subject
                    else:               # не осталось - ищем значение после типа субъекта
                        try:
                            subject = self.field_home[breaks_sorted[break_sorted] + len(breaks_name[break_sorted]) + 1
                                                           :breaks_sorted[list(breaks_sorted.keys())[i + 1]]]
                        except IndexError:
                            subject = self.field_home[breaks_sorted[break_sorted] + len(breaks_name[break_sorted]) + 1:]
                        for j in range(5):
                            for split_field in SPLIT_FIELDS:
                                subject = subject.strip(split_field)
                        if len(subject):
                            output_dict[break_sorted] = subject
                            breaks_len[break_sorted] = len(subject)
                    break_sorted_last = break_sorted
                    if break_sorted == 11:  # дом - нет типа субъекта
                        pass
                    elif break_sorted == 12:  # корпус - нет типа субъекта
                        pass
                    elif break_sorted == 13:  # квартира - нет типа субъекта
                        pass
                    else: # у остальных есть тип субъекта
                        pass


                # записываем пустые строки если не нашли такую позицию
                for i in range(13):
                    if i in output_dict.keys():
                        output_list.append(output_dict[i])
                    else:
                        output_list.append('')

                pass



tek_adress = ' 322236 Самарская обл, г Самара, Балаковская ул, дом № 20, кв.113;  Самарская обл, г Самара, Балаковская ул, дом № 20, кв.113'

cut_adress = 'Самарская Самара Балаковская'
res = requests.get('http://127.0.0.1:23332/find/' + cut_adress + '?strong=1')
if res.status_code == 200:
    ajson = json.loads(bytes.decode(res.content))
    print(ajson[0]['text'])
# добавить исключение из текста сокращений ФИАС и обрезка хвоста, начиная от дома
adress_reg = FullAdress(tek_adress)
a = adress_reg.get_values()
pass
