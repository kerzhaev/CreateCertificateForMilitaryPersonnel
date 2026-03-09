Attribute VB_Name = "UDFs_FIO"


Private Function IsMan(ByVal sName As String) As Boolean
    arMenNames = Array("Абай", "Абрам", "Абраам", "Аваз", "Авазбек", "Авдей", "Адилет", "Адольф", "Азамат", "Акбар", "Аксентий", "Агафон", "Айбек", "Айрат", "Алдар", "Алишер", "Алан", "Александр", "Алексей", "Али", "Алмат", "Альберт", "Альвиан", "Альфред", "Анатолий", "Андрей", "Антон", "Антонин", "Аристарх", "Аркадий", "Армен", "Арнольд", "Арон", "Арсен", "Арсений", "Артем", "Артём", "Артемий", _
        "Артур", "Аскольд", "Афанасий", "Ашот", "Батыр", "Бауыржан", "Богдан", "Борис", "Вадим", "Валентин", "Валерий", "Валерьян", "Варлам", "Василий", "Вахтанг", "Венедикт", "Вениамин", "Виктор", "Виталий", "Влад", "Владилен", "Владимир", "Владислав", "Владлен", "Вольф", "Всеволод", "Вячеслав", "Гавриил", "Гаврил", "Гайдар", _
        "Геласий", "Геннадий", "Генрих", "Георгий", "Герасим", "Герман", "Глеб", "Гордей", "Григорий", "Гурген", "Давид", "Дамир", "Даниил", "Данил", "Данияр", "Дастан", "Демьян", "Денис", "Диас", "Динишбек", "Дмитрий", "Дорофей", "Евгений", "Евграф", "Евдоким", "Евсей", "Егор", _
        "Еремей", "Ернар", "Ермолай", "Ефим", "Жонибек", "Заур", "Зиновий", "Иакинф", "Иван", "Игнатий", "Игнат", "Игорь", "Иларион", "Илларион", "Ильдар", "Ильшат", "Илья", "Иннокентий", "Иосиф", "Ипполит", "Ирек", "Ириней", "Исидор", "Исаак", "Исхак", "Иулиан", "Казимир", "Кайрат", "Камиль", "Карл", "Касьян", "Керим", "Ким", "Кирилл", "Клавдий", "Кондрат", "Константин", _
        "Кристиан", "Кузьма", "Куприян", "Лаврентий", "Лев", "Ленар", "Леонард", "Леонид", "Леонтий", "Лука", "Лукий", "Лукьян", "Людвиг", "Магомед", "Магомет", "Майк", "Макар", "Максат", "Макс", "Максим", "Марат", "Марк", "Мартын", "Матвей", "Махач", "Махмуд", "Мелентий", "Мирлан", "Мирослав", _
        "Митрофан", "Михаил", "Модест", "Моисей", "Мстислав", "Мурад", "Мухамед", "Мухаммед", "Муса", "Мэлор", "Наум", "Никита", "Никифор", "Николай", "Нурбек", "Нуржан", "Нурлан", "Олег", "Онисим", "Осип", "Отар", "Павел", "Пантелеймон", "Парфений", "Пётр", "Петр", "Платон", "Порфирий", "Прокопий", "Протасий", "Прохор", "Радомир", "Разумник", "Рамазан", "Рамзан", "Рафаэль", _
        "Рафик", "Ринат", "Роман", "Роберт", "Ростислав", "Рубен", "Рудольф", "Руслан", "Рустам", "Рустем", "Сабир", "Савва", "Савелий", "Святослав", "Семён", "Семен", "Серафим", "Сергей", "Серик", "Созон", "Соломон", "Спиридон", "Станислав", "Степан", "Султан", "Тагир", "Тарас", "Темир", "Темирхан", "Тигран", "Тимофей", "Тимур", "Тихон", "Трифон", _
        "Трофим", "Фадей", "Фаддей", "Федор", "Фёдор", "Федосей", "Федот", "Феликс", "Филат", "Филипп", "Фома", "Фрол", "Харитон", "Хафиз", "Христофор", "Чеслав", "Шамиль", "Шамхал", "Эдуард", "Эльдар", "Эльман", "Эмиль", "Эммануил", "Эраст", "Юлиан", "Юлиус", "Юлий", "Юрий", "Юстин", "Яков", "Якун", "Ян", "Ярослав")

    For i = LBound(arMenNames) To UBound(arMenNames)
        If sName = arMenNames(i) Then
            IsMan = True
            Exit Function
        End If
    Next i


End Function

Private Function IsWoman(ByVal sName As String) As Boolean

    arWomenNames = Array("Августа", "Авдотья", "Агафья", "Агриппина", "Адиля", "Аида", "Аиша", "Айару", "Айгерим", "Айгуль", "Айнур", "Айнура", "Аксинья", "Акулина", "Алевтина", "Александра", "Александрина", "Алексина", "Алёна", "Алеся", "Алина", "Алиса", "Алла", "Алсу", "Алтынай", "Альбина", "Альфия", "Амина", "Амра", "Анастасия", "Ангелина", _
        "Анель", "Анжела", "Анжелика", "Анна", "Антонина", "Арина", "Армине", "Аружан", "Асель", "Асем", "Асмик", "Асоль", "Ася", "Аурика", "Ая", "Аяла", "Айя", "Белла", "Бэлла", "Бося", "Валентина", "Валерия", "Варвара", "Василиса", "Вера", "Вероника", "Виктория", "Виолетта", "Владилена", "Владислава", "Галина", "Глафира", "Гузель", "Гулнар", "Гульнара", _
        "Гульшат", "Гюзель", "Давлят", "Дана", "Дарья", "Дария", "Джамиля", "Диана", "Диляра", "Дина", "Динара", "Ева", "Евгения", "Евдокия", "Евпраксия", "Евфросиния", "Екатерина", "Елена", "Елизавета", "Жанат", "Жанар", "Жанара", "Жанна", "Жанетта", "Жулдыз", "Зауре", "Земфира", "Зимфира", "Зинаида", "Злата", _
        "Зоя", "Иванна", "Инга", "Инесса", "Инна", "Ираида", "Ирина", "Ирма", "Ия", "Капитолина", "Карина", "Каринэ", "Каролина", "Катерина", "Катрин", "Кира", "Клавдия", "Клара", "Кристина", "Ксения", "Лада", "Лариса", "Лейла", "Лейли", "Лейсан", "Лениза", "Леся", "Лиана", "Лига", "Лидия", _
        "Лилия", "Лия", "Лэйсэн", "Любовь", "Людмила", "Ляйсан", "Мадина", "Майя", "Маргарита", "Маржан", "Мариана", "Марианна", "Марина", "Мария", "Марфа", "Матрёна", "Матрена", "Мацак", "Милена", "Милана", "Мира", "Мирослава", "Муза", "Муит", "Надежда", "Назира", "Наида", "Наина", "Наринэ", "Наталья", "Наталия", "Нелли", "Нигина", "Николета", _
        "Нина", "Нинель", "Нонна", "Оксана", "Октябрина", "Олеся", "Ольга", "Пелагея", "Полина", "Прасковья", "Раиса", "Регина", "Ригина", "Римма", "Рита", "Роза", "Розалия", "Ромина", "Русина", "Руслана", "Руфина", "Сабина", "Салтанат", "Светлана", "Серафима", "Снежана", "София", "Софья", "Стелла", "Стефания", _
        "Таисия", "Тайя", "Тамара", "Татевик", "Татьяна", "Томирис", "Ульяна", "Фаина", "Феврония", "Фёкла", "Феодора", "Ханзада", "Целестина", "Шамиля", "Элеонора", "Элина", "Элла", "Эльвира", "Эльза", "Эмилия", "Эмма", "Эсфирь", "Юлия", "Яна", "Ярослава")
    
    For i = LBound(arWomenNames) To UBound(arWomenNames)
        If sName = arWomenNames(i) Then
            IsWoman = True
            Exit Function
        End If
    Next i

End Function

Function FIO(NameAsText As String, Optional NameCase As String = "И", Optional ShortForm As Boolean = False) As String
Attribute FIO.VB_Description = "Возвращает ФИО в правильной последовательности и заданном падеже для исходного имени в ячейке."
    'выстраивает ФИО в правильном порядке, склоняет по падежам и, при желании, выводит в сокращенной форме

    Dim iGender As Integer
    Dim sName$, sName2$, sMidName$, sMidName2$, sSurName$, sSurName2$
    Dim arWords
    
    '----------------------- ОПРЕДЕЛЯЕМ ГДЕ ИМЯ, ГДЕ ФАМИЛИЯ, А ГДЕ ОТЧЕСТВО -----------------------------------------
    iGender = 0
    iGender = GetSex(NameAsText)        'определяем пол
    arWords = Split(WorksheetFunction.Trim(NameAsText), " ")        'разбираем ФИО на слова
        
    'если в ячейке полное ФИО, т.е. есть и отчество
    If UBound(arWords) = 2 Then
        If iGender = -1 Then
            If Right(arWords(1), 3) = "вич" Or Right(arWords(1), 3) = "тич" Then
                sSurName = arWords(2)
                sName = arWords(0)
                sMidName = arWords(1)
            End If
            If Right(arWords(2), 3) = "вич" Or Right(arWords(2), 3) = "тич" Then
                sSurName = arWords(0)
                sName = arWords(1)
                sMidName = arWords(2)
            End If
        End If

        If iGender = 1 Then
            If Right(arWords(1), 3) = "вна" Or Right(arWords(1), 3) = "чна" Then
                sSurName = arWords(2)
                sName = arWords(0)
                sMidName = arWords(1)
            End If
            If Right(arWords(2), 3) = "вна" Or Right(arWords(2), 3) = "чна" Then
                sSurName = arWords(0)
                sName = arWords(1)
                sMidName = arWords(2)
            End If
        End If
    End If
        
    'если есть только фамилия и имя без отчества - ищем имена по справочникам
    If UBound(arWords) = 1 Then
        If IsMan(arWords(0)) Or IsWoman(arWords(0)) Then
            sName = arWords(0)
            sSurName = arWords(1)
        End If
        If IsMan(arWords(1)) Or IsWoman(arWords(1)) Then
            sName = arWords(1)
            sSurName = arWords(0)
        End If
    End If
    
    'если в ячейке только одно слово - пытаемся по полу определить - это имя или фамилия
    If UBound(arWords) = 0 Then
        If IsMan(arWords(0)) Or IsWoman(arWords(0)) Then
            'если пол определился - значит это имя
            sName = arWords(0)
        Else
            'если не определился - значит это фамилия
            sSurName = arWords(0)
            'пытаемся определить пол по окончанию фамилии, если возможно
            If sSurName Like "*ов" Or sSurName Like "*ев" Or sSurName Like "*ин" Or sSurName Like "*ий" Or sSurName Like "*ой" Then iGender = -1
            If sSurName Like "*ва" Or sSurName Like "*на" Or sSurName Like "*ая" Then iGender = -1
            'если пол так и не определился, то выходим
            If iGender = 0 Then
                FIO = ""
                Exit Function
            End If
        End If
    End If

    
    
    '--------------------------- ИМЕНИТЕЛЬНЫЙ ПАДЕЖ (КТО) ---------------------------------------------------------
    sName2 = sName
    sSurName2 = sSurName
    sMidName2 = sMidName
        
    '--------------------------- ДАТЕЛЬНЫЙ ПАДЕЖ (КОМУ) ---------------------------------------------------------
    
    If UCase(NameCase) = "Д" Or UCase(NameCase) = "D" Then
        'формируем дательный падеж для имени
        If sName <> "" Then
            sName2 = sName
            If iGender = -1 Then
                If sName Like "*[ая]" Then sName2 = Left(sName, Len(sName) - 1) & "е"
                If sName Like "*[бвгджзклмнпрстфхцчшщ]" Then sName2 = sName & "у"
                If sName Like "*[йь]" Then sName2 = Left(sName, Len(sName) - 1) & "ю"
            End If
            If iGender = 1 Then
                If sName Like "*а" Then sName2 = Left(sName, Len(sName) - 1) & "е"
                If sName Like "*ия" Then sName2 = Left(sName, Len(sName) - 1) & "и"
                If sName Like "*ея" Then sName2 = Left(sName, Len(sName) - 1) & "и"
                If sName Like "*ья" Then sName2 = Left(sName, Len(sName) - 1) & "е"
                If sName Like "*ь" Then sName2 = Left(sName, Len(sName) - 1) & "и"
            End If
        End If
        
        'формируем дательный падеж для отчества
        If sMidName <> "" Then
            sMidName2 = sMidName
            If Right(sMidName, 1) = "а" Then sMidName2 = Left(sMidName, Len(sMidName) - 1) & "е"
            If Right(sMidName, 1) = "ч" Then sMidName2 = sMidName & "у"
        End If
        
        'формируем дательный падеж для фамилии
        If sSurName <> "" Then
            sSurName2 = sSurName
            If iGender = -1 Then
                If sSurName Like "*а" Then sSurName2 = Left(sSurName, Len(sSurName) - 1) & "е"
                If sSurName Like "*й" Then sSurName2 = Left(sSurName, Len(sSurName) - 2) & "ому"
                If sSurName Like "*ай" Then sSurName2 = Left(sSurName, Len(sSurName) - 1) & "ю"
                If sSurName Like "*ь" Then sSurName2 = Left(sSurName, Len(sSurName) - 1) & "ю"
                If sSurName Like "*[бвгджзклмнпрстфхцчшщ]" Then sSurName2 = sSurName & "у"
                If sSurName Like "*ых" Or sSurName Like "*их" Or sSurName Like "*иа" Or sSurName Like "*ия" Or sSurName Like "*уя" Or sSurName Like "*ая" Then sSurName2 = sSurName
                If sSurName Like "*ок" Or sSurName Like "*их" Then sSurName2 = Left(sSurName, Len(sSurName) - 2) & "ку"
            End If
            If iGender = 1 Then
                If sSurName Like "*а" Then sSurName2 = Left(sSurName, Len(sSurName) - 1) & "ой"
                If sSurName Like "*ая" Then sSurName2 = Left(sSurName, Len(sSurName) - 2) & "ой"
                If sSurName Like "*[бвгджзклмнпрстфхцчшщ]" Then sSurName2 = sSurName
            End If
            
        End If
    End If

        
    '--------------------------- РОДИТЕЛЬНЫЙ ПАДЕЖ (КОГО) ---------------------------------------------------------
    
    If UCase(NameCase) = "Р" Or UCase(NameCase) = "R" Then
        'формируем родительный падеж для имени
        If sName <> "" Then
            sName2 = sName
            If iGender = -1 Then
                If sName Like "*а" Then sName2 = Left(sName, Len(sName) - 1) & "ы"
                If sName Like "*[бвгджзклмнпрстфхцчшщ]" Then sName2 = sName & "а"
                If sName Like "*[йь]" Then sName2 = Left(sName, Len(sName) - 1) & "я"
            End If
            If iGender = 1 Then
                If sName Like "*а" Then sName2 = Left(sName, Len(sName) - 1) & "ы"
                If sName Like "*ия" Then sName2 = Left(sName, Len(sName) - 1) & "и"
                If sName Like "*ея" Then sName2 = Left(sName, Len(sName) - 1) & "и"
                If sName Like "*ья" Then sName2 = Left(sName, Len(sName) - 1) & "и"
                If sName Like "*ь" Then sName2 = Left(sName, Len(sName) - 1) & "и"
            End If
        End If
        
        'формируем родительный падеж для отчества
        If sMidName <> "" Then
            sMidName2 = sMidName
            If Right(sMidName, 1) = "а" Then sMidName2 = Left(sMidName, Len(sMidName) - 1) & "ы"
            If Right(sMidName, 1) = "ч" Then sMidName2 = sMidName & "а"
        End If
        
        'формируем родительный падеж для фамилии
        If sSurName <> "" Then
            sSurName2 = sSurName
            If iGender = -1 Then
                If sSurName Like "*а" Then sSurName2 = Left(sSurName, Len(sSurName) - 1) & "ы"
                If sSurName Like "*й" Then sSurName2 = Left(sSurName, Len(sSurName) - 2) & "ого"
                If sSurName Like "*ай" Then sSurName2 = Left(sSurName, Len(sSurName) - 1) & "я"
                If sSurName Like "*ь" Then sSurName2 = Left(sSurName, Len(sSurName) - 1) & "я"
                If sSurName Like "*[бвгджзклмнпрстфхцчшщ]" Then sSurName2 = sSurName & "а"
                If sSurName Like "*ок" Then sSurName2 = Left(sSurName, Len(sSurName) - 2) & "ка"
                If sSurName Like "*ых" Or sSurName Like "*их" Or sSurName Like "*иа" Or sSurName Like "*ия" Or sSurName Like "*уя" Or sSurName Like "*ая" Then sSurName2 = sSurName
            End If
            If iGender = 1 Then
                If sSurName Like "*а" Then sSurName2 = Left(sSurName, Len(sSurName) - 1) & "ой"
                If sSurName Like "*ая" Then sSurName2 = Left(sSurName, Len(sSurName) - 2) & "ой"
                If sSurName Like "*[бвгджзклмнпрстфхцчшщ]" Then sSurName2 = sSurName
            End If
        End If
    End If
    
    '------------------- ВЫВОДИМ РЕЗУЛЬТАТЫ -------------------------------------------------------------------------
    If ShortForm Then
        FIO = sSurName2 & " " & Left(sName2, 1) & "." & Left(sMidName2, 1) & "."
    Else
        FIO = sSurName2 & " " & sName2 & " " & sMidName2
    End If
            
    If sMidName = "" Then FIO = Left(FIO, Len(FIO) - 1)    'если нет отчества - убираем лишний последний пробел или точку
    FIO = Trim(FIO)
        
End Function


Function GetSex(ByVal cell As String) As Integer
    Dim arWords
    Dim iGender As String
    Dim i As Integer

    'разбираем на слова
    arWords = Split(WorksheetFunction.Trim(cell), " ")
    
    iGender = 0
    'если имя есть в справочниках - определяем пол сразу
    For i = LBound(arWords) To UBound(arWords)
        If IsMan(arWords(i)) Then iGender = "-1"
        If IsWoman(arWords(i)) Then iGender = 1
    Next i
        
    'если имени нет в справочниках - пытаемся определить по отчеству
    If iGender = 0 Then
        For i = LBound(arWords) To UBound(arWords)
            If Right(arWords(i), 3) = "вна" Or Right(arWords(i), 3) = "чна" Then iGender = 1
            If Right(arWords(i), 3) = "вич" Or Right(arWords(i), 3) = "ьич" Then iGender = -1
        Next i
    End If
    
    GetSex = iGender
 
End Function







