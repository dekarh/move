# -*- coding: utf-8 -*-

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
SPLIT_FIELD = ','
#SPLIT_FIELD = '_x0003_'  # Разделитель для адреса в одной строке (бывает '_x0003_')

#ORDER_FIELD = [13, 0, 8, 1, 9, 2, 10, 3, 11, 4, 12, 5, 6, 7] # См def FullAddress(get_values():
# ['Индекс', 'Регион', 'Тип_региона', 'Район', 'Тип_района', 'Город', 'Тип_города',
#  'Населенный_пункт', 'Тип_населенного_пункта', 'Улица', 'Тип_улицы', 'Дом', 'Корпус', 'Квартира']

REG_TYPES = ['обл', 'о', 'область', 'респ', 'республика', 'край', 'кр', 'ар', 'ао', 'авт']

DISTRICT_TYPES = ['р-н', 'р', 'район']

CITY_TYPES = ['г', 'гор', 'город']

NP_TYPES = ['пгт', 'пос', 'поселение', 'поселок', 'посёлок', 'п', 'рп', 'кп', 'к', 'пс', 'сс', 'смн', 'вл', 'дп',
            'нп', 'пст', 'ж/д_ст', 'с', 'м', 'д', 'дер', 'сл', 'ст', 'ст-ца', 'х', 'рзд', 'у', 'клх', 'свх', 'зим', 'мкр']

STREET_TYPES = ['аллея', 'а', 'бульвар', 'б-р', 'в/ч', 'городок', 'гск', 'кв-л', 'линия', 'наб', 'пер', 'переезд', 'пл',
                'пр-кт', 'проезд', 'тер', 'туп', 'ул', 'ш', ]

HOUSE_CUT_NAME = ['дом', 'д']
CORPUS_CUT_NAME = ['корп', 'корпус']
APARTMENT_CUT_NAME = ['кв']
########################################################################################################################
# ЗНАЧЕНИЕ В ПОЛЕ "ПОЛ" В ИСХОДНОМ ФАЙЛЕ
#FEMALE_GENDER_VALUE = 'Ж'
#MALE_GENDER_VALUE = 'М'
FEMALE_GENDER_VALUE = 'Женский'
MALE_GENDER_VALUE = 'Мужской'
#FEMALE_GENDER_VALUE = 'Жен.'
#MALE_GENDER_VALUE = 'Муж.'
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

FIO_LABELS = ['Фамилия', 'Имя', 'Отчество']

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

import string
import re


class BaseClass:

    def __setattr__(self, name, value):
        if isinstance(value, (int, str)):
            self.__dict__[name] = str(value).strip()
        else:
            self.__dict__[name] = value


def normalize(*args):
    result = []
    for arg in args:
        if arg == NULL_VALUE:
            result.append(NEW_NULL_VALUE)
        else:
            result.append(str(arg).strip())
    return result


def normalize_snils(snils):
    snils = str(snils).strip()
    snilsX = ''
    if snils != NULL_VALUE and snils != '' and isinstance(snils, str):
        try:
            for cc in snils:
                if cc in string.digits:
                    snilsX = snilsX+cc
            if len(snilsX) < LEN_SNILS:
                snilsX = '0' * (LEN_SNILS - len(snilsX)) + snilsX
            elif len(snilsX) == LEN_SNILS:
                pass
            else:
                return ERROR_VALUE
            return snilsX
        except TypeError:
            return ERROR_VALUE
    else:
        return ERROR_VALUE


def field2fio(field):
    if len(field) > 0 and field != NULL_VALUE:
        first_name, second_name, third_name = '', '', ''
        field = field.strip().replace('  ',' ').replace('  ',' ').split(' ')
        for i, word in enumerate(field):
            if i == 0:
                first_name = field[i]
            elif i == 1:
                second_name = field[i]
            else:
                third_name += field[i] + ' '
        if len(third_name) > 0:
            third_name = third_name[:-1]
        return first_name, second_name, third_name
    else:
        return NEW_NULL_VALUE_FOR_ALL_TEXT

def field2addr(field):
    addr_name, addr_type = '', ''
    if len(field) > 0 and field != NULL_VALUE:
        new_field = ''
        for i, ch in enumerate(field):          # убираем точки и запятые
            if ch == '.' or ch == ',':
                ch = ''
            new_field = new_field + ch
        field = new_field.strip().split(' ')
        TYPES = [REG_TYPES, DISTRICT_TYPES, CITY_TYPES, NP_TYPES, STREET_TYPES]
        for i, word in enumerate(field):
            addr_type_vrem = ''
            for l in TYPES:
                for ll in l:
                    if word.lower() == ll.lower():
                        addr_type_vrem = ll
            if addr_type_vrem == '':
                addr_name = addr_name + ' ' + word
            else:
                addr_type = addr_type_vrem
    return addr_name, addr_type

#class Gender(BaseClass):
#    def __init__(self, third_name='', gender_field_exists=False, gender=''):
#        self.female_gender_value = FEMALE_GENDER_VALUE
#        self.male_gender_value = MALE_GENDER_VALUE
#        self.third_name = str(third_name).strip()
#        self.gender_field_exists = gender_field_exists
#        self.gender = gender.strip()

#    def gender_from_fio(self):
#        if self.third_name == '':
#            return ERROR_VALUE
#        third_name = self.third_name
#        third_name = third_name.split(' ')
#        if len(third_name) == 1:
#            if ''.join(third_name[0][-3:]).lower() == 'вна':
#                gender = '0'
#            elif ''.join(third_name[0][-3:]).lower() == 'вич':
#                gender = '1'
#            else:
#                gender = ERROR_VALUE
#        elif len(third_name) == 2:
#            if third_name[-1].lower() in EAST_GENDER[0]:  # женщина
#                gender = '0'
#            elif third_name[-1].lower() in EAST_GENDER[1]:  # мужчина
#                gender = '1'
#            else:
#                gender = ERROR_VALUE
#        else:
#            gender = ERROR_VALUE
#        return gender

#    def set_gender_value(self, male_value, female_value):
#        self.female_gender_value = female_value.lower()
#        self.male_gender_value = male_value.lower()

#    def get_gender_value(self):
#        return self.female_gender_value, self.male_gender_value

#    def normalize_gender(self):
#        gender = self.gender
#        gender = gender.lower()
#        if gender == self.female_gender_value:
#            return '0'
#        elif gender == self.male_gender_value:
#            return '1'
#        else:
#            return self.gender_from_fio()

#    def get_value(self):
#        if self.gender_field_exists:
#            return self.normalize_gender()
#        else:
#            return self.gender_from_fio()

def normalize_gender(gender):
    gender = str(gender).strip()
    if gender =='':
        return NEW_NULL_VALUE_FOR_GENDER
    elif len(gender) > 1 and (gender.strip()!=FEMALE_GENDER_VALUE and gender.strip()!=MALE_GENDER_VALUE):
        return NEW_NULL_VALUE_FOR_GENDER
    else:
        if gender.strip() == FEMALE_GENDER_VALUE:
            return '1'
        else:
            return '0'

def normalize_text(tx):
    tx = str(tx).strip()
    if len(tx) <= 1:
        return NEW_NULL_VALUE_FOR_ALL_TEXT
    else:
        return tx

def normalize_date(date):
    date = str(date)
    result = re.findall(r'\b(\d{4}|\d{2})[\.:-](\d{2})[\.:-](\d{4}|\d{2})\b', date)
    if len(result) > 0:
        if result[0] == NULL_VALUE:
            return NEW_NULL_VALUE_FOR_DATE
        if len(result[0][0]) == 4:
            result[0] = result[0][::-1]
        elif len(result[0][2]) == 2:
            if result[0][2] < 20:
                result[0][2] = '20' + result[0][2]
            elif result[0][2] > 20:
                result[0][2] = '19' + result[0][2]
        return '.'.join(result[0])
    else:
        return NEW_NULL_VALUE_FOR_DATE


# print(normalize_date('01.09.2003'))

# normalize место рождения


class Passport(BaseClass):
    def __init__(self, seriya='', nomer='', date='', who='', cod=''):
        self.seriya = str(seriya).strip()
        self.nomer = str(nomer).strip()
        self.date = normalize_date(date)
        self.who = normalize_text(who)
        self.cod = str(cod).strip()

    def __setattr__(self, name, value):
        if name == 'date':
            self.__dict__[name] = normalize_date(value)
        else:
            if isinstance(value, (int, str)):
                self.__dict__[name] = str(value)
            else:
                self.__dict__[name] = value

    def normalize_seriya(self):
        if self.seriya != NULL_VALUE and self.seriya != '' and isinstance(self.seriya, str):
            try:
                self.seriya = ''.join([char for char in self.seriya if char in string.digits])
                if len(self.seriya) == 3:
                    self.seriya = '0' + self.seriya
                elif len(self.seriya) == 4:
                    pass
                else:
                    return NEW_NULL_VALUE_FOR_SERIYA_PASSPORTA
#                   return ERROR_VALUE
#                self.seriya = self.seriya[:2] + ' ' + self.seriya[2:]         # Тот самый пробел между ## ##
                return self.seriya
            except TypeError:
                return NEW_NULL_VALUE_FOR_SERIYA_PASSPORTA
        else:
            return NEW_NULL_VALUE_FOR_SERIYA_PASSPORTA

    def normalize_nomer(self):
        if self.nomer != NULL_VALUE and self.nomer != '' and isinstance(self.seriya, str):
            try:
                self.nomer = ''.join([char for char in self.nomer if char in string.digits])
                if len(self.nomer) < LEN_PASSPORT_NOMER and len(self.nomer) > 0 :
                    self.nomer = '0' * (LEN_PASSPORT_NOMER - len(self.nomer)) + self.nomer
                elif len(self.nomer) == LEN_PASSPORT_NOMER:
                    pass
                else:
                    return NEW_NULL_VALUE_FOR_NOMER_PASSPORTA
#                    return ERROR_VALUE
                return self.nomer
            except TypeError:
                return NEW_NULL_VALUE_FOR_NOMER_PASSPORTA
        else:
            return NEW_NULL_VALUE_FOR_NOMER_PASSPORTA


    def normalize_who(self):
        if self.who == NULL_VALUE:
            self.who = NEW_NULL_VALUE
        return self.who

    def normalize_cod(self):
        if self.cod != NULL_VALUE and self.cod != '':
            self.cod = ''.join([char for char in self.cod if char in string.digits])
            if len(self.cod) < LEN_PASSPORT_COD and len(self.cod) > 0:
                self.cod = '0' * (LEN_PASSPORT_COD - len(self.cod)) + self.cod
            elif len(self.cod) == LEN_PASSPORT_COD:
                pass
            else:
                return NEW_NULL_VALUE_FOR_COD_PASSPORTA
#                return ERROR_VALUE
            self.cod = self.cod[:3] + '-' + self.cod[3:]
            return self.cod
        else:
            return NEW_NULL_VALUE_FOR_COD_PASSPORTA

    def get_values(self):
        return self.normalize_seriya(), self.normalize_nomer(), self.date, self.normalize_who(), self.normalize_cod()


def normalize_index(index):
    index = str(index).strip()
    if index != NULL_VALUE and index != '' and isinstance(index, str):
        try:
            index = ''.join([char for char in index if char in string.digits])
            if len(index) < LEN_INDEX_NOMER:
                index = '0' * (LEN_INDEX_NOMER - len(index)) + index
            elif len(index) == LEN_INDEX_NOMER:
                pass
            else:
                return NEW_NULL_VALUE_FOR_INDEX
#                return ERROR_VALUE
            return index
        except TypeError:
            return NEW_NULL_VALUE_FOR_INDEX
    else:
        return NEW_NULL_VALUE_FOR_INDEX

def normalize_home(tx):
        tx = str(tx).strip()
        numbers = True
        for i in range(len(tx)):
            if tx[i] not in string.digits:
                numbers = False
        if len(tx) < 1:
            return tx
        elif len(tx) > 10:
            return NEW_NULL_VALUE_FOR_HOME
        elif numbers:
            if int(tx) > 1500:
                return NEW_NULL_VALUE_FOR_HOME
            else:
                return tx
        else:
            return tx


class FullAdress(BaseClass):
    def __init__(self, field=''):
        self.field = str(field)
        self.full_adress = []
        self.FULL_ADRESS_DICT = {}
        for label in FULL_ADRESS_LABELS:
            self.FULL_ADRESS_DICT[label] = ''
        self.iter_types = [DISTRICT_TYPES, CITY_TYPES, NP_TYPES, STREET_TYPES, HOUSE_CUT_NAME, CORPUS_CUT_NAME, APARTMENT_CUT_NAME]

    def normalize_adress(self):
        if len(self.field) != 0 and self.field != NULL_VALUE:
            self.field = self.field.lower()
            values = self.field.split(SPLIT_FIELD)
            for i, word in enumerate(values):
                n = []
                word = word.strip()
                if i == 0:
                    n = [char for char in word if char in string.digits]
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

    def get_values(self):  # Когда адрес 414000, г. Астрахань, ул. Такая, д. Т...
        output_list = []
        for elem in self.create_output_list():
            output_list.append(elem.strip())
        return output_list

    a = """
    def get_values(self):                           # Когда все поля по раздельности...
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


    # def __call__(self, *args, **kwargs):
    #     return self.create_output_list()


# f = FullAdress('123592, Москва г, строгинский бульвар, д. 26, корпус 2, кв. 425')
# print(f.get_values())
    """

class Phone(BaseClass):
    def __init__(self, tel_mob='', tel_rod='', tel_dom=''):
        self.tel_mob = str(tel_mob).strip()
        self.tel_rod = str(tel_rod).strip()
        self.tel_dom = str(tel_dom).strip()

    def normalize_tel_number(self, tel):
        tel = tel.strip()
        if tel == '' or tel == NULL_VALUE:
            return ERROR_VALUE
        tel = str(tel).strip()
        tel = ''.join([char for char in tel if char in string.digits])
        if len(tel) == 11:
            if tel[0] in ['7', '8', '9']:
                tel = '7' + tel[1:]
            else:
                return ERROR_VALUE
        elif len(tel) == 10:
            tel = '7' + tel
        else:
            return ERROR_VALUE
        return tel

    def poryadoc(self, *tels):
        tels = sorted(tels)
#        tels.reverse()
        return list(tels)

    def get_values(self):
        self.tel_mob = self.normalize_tel_number(self.tel_mob)
        self.tel_rod = self.normalize_tel_number(self.tel_rod)
        self.tel_dom = self.normalize_tel_number(self.tel_dom)
#        if self.tel_rod == self.tel_mob:
#            self.tel_rod = ''
#        self.tel_dom = self.normalize_tel_number(self.tel_dom)
#        if self.tel_dom == self.tel_mob or self.tel_dom == self.tel_rod:
#            self.tel_rod = ''
        return self.poryadoc(self.tel_mob, self.tel_rod, self.tel_dom)

# p = Phone()
# p.tel_dom = 89040964007
# p.tel_rod = 89257349331
# p.tel_mob = 'dd'
# print(p.get_values())

def intl(a):               # белиберду в цифры или 0
    try:
        if a != None:
            a = str(a).strip()
            if  a != '':
                a = ''.join([char for char in a if char in string.digits])
                if len(a) > 0:
                    return int(a)
                else:
                    return 0
        return 0
    except TypeError:
        return 0

