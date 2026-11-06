"use strict";
/* Факты дня про кока-колу. Каждый факт сверен с источником, ссылка показывается под фактом.
   Чтобы добавить факт, допишите строку { t: "текст", u: "ссылка на источник" } в конец списка. */

const FACTS = [
  { t: "Первый стакан кока-колы налили 8 мая 1886 года в аптеке Jacobs' Pharmacy в Атланте. Стоил он 5 центов.", u: "https://www.coca-colacompany.com/about-us/history/the-birth-of-a-refreshing-idea" },
  { t: "Кока-колу изобрёл атлантский фармацевт Джон Стит Пембертон. Он сам принёс в аптеку кувшин сиропа, который там смешали с газированной водой.", u: "https://www.coca-colacompany.com/about-us/history/the-birth-of-a-refreshing-idea" },
  { t: "В первый год продавали в среднем всего 9 стаканов кока-колы в день.", u: "https://www.coca-colacompany.com/about-us/history/the-birth-of-a-refreshing-idea" },
  { t: "Сегодня напитки компании Coca-Cola пьют около 2,2 миллиарда раз в день более чем в 200 странах.", u: "https://www.coca-colacompany.com/about-us/coca-cola-system" },
  { t: "Название придумал бухгалтер и партнёр Пембертона Фрэнк Робинсон: он решил, что две заглавные «C» будут хорошо смотреться в рекламе.", u: "https://www.coca-colacompany.com/about-us/history/the-birth-of-a-refreshing-idea" },
  { t: "Знаменитый логотип тоже написал Фрэнк Робинсон — от руки, каллиграфическим почерком в стиле «спенсериан».", u: "https://www.coca-cola.com/np/en/about-us/history/1880-decade" },
  { t: "Первая газетная реклама вышла в The Atlanta Journal и называла кока-колу «восхитительным и освежающим напитком».", u: "https://www.coca-cola.com/np/en/about-us/history/1880-decade" },
  { t: "Уже в 1887 году кока-колу продвигали купонами на бесплатный стакан.", u: "https://www.coca-cola.com/np/en/about-us/history/1880-decade" },
  { t: "Полный контроль над Coca-Cola обошёлся предпринимателю Аса Кэндлеру всего в 2300 долларов: права он выкупал по частям у Пембертона и его партнёров.", u: "https://coca-colacompany.com/about-us/history/the-asa-candler-era" },
  { t: "Товарный знак «Coca-Cola» зарегистрировали в США 31 января 1893 года, хотя напиток продавали под этим именем с 1886-го.", u: "https://coca-colacompany.com/about-us/history/the-asa-candler-era" },
  { t: "Короткое «Coke» стало зарегистрированной торговой маркой только в 1945 году.", u: "https://coca-colacompany.com/about-us/history/the-asa-candler-era" },
  { t: "Первым кока-колу в бутылки разлил в 1894 году Джозеф Биденхарн: он поставил разливочную машину прямо у себя в магазине в Виксберге, штат Миссисипи.", u: "https://coca-colacompany.com/about-us/history/the-asa-candler-era" },
  { t: "В 1899 году два юриста из Чаттануги получили права на розлив кока-колы почти по всем США. За следующие 20 лет число разливочных заводов выросло с двух до более чем тысячи.", u: "https://coca-colacompany.com/about-us/history/the-asa-candler-era" },
  { t: "Уже в 1895 году Аса Кэндлер писал в годовом отчёте, что кока-колу пьют «в каждом штате и территории США».", u: "https://coca-colacompany.com/about-us/history/the-asa-candler-era" },
  { t: "Первый завод сиропа за пределами Атланты открыли в 1894 году в Далласе, а в 1895-м — в Чикаго и Лос-Анджелесе.", u: "https://coca-colacompany.com/about-us/history/the-asa-candler-era" },
  { t: "В 1915 году объявили конкурс на бутылку, которую можно узнать на ощупь в темноте или по осколку на земле. На разработку выделили до 500 долларов.", u: "https://www.coca-colacompany.com/about-us/history/the-history-of-the-coca-cola-contour-bottle" },
  { t: "Фигурную бутылку придумали на заводе Root Glass в Терре-Хоте, штат Индиана. Вдохновились рисунком какао-боба — вытянутого и с рёбрами.", u: "https://www.coca-colacompany.com/about-us/history/the-history-of-the-coca-cola-contour-bottle" },
  { t: "Патент на фигурную бутылку выдали 16 ноября 1915 года на имя мастера завода, шведского иммигранта Александра Самуэльсона.", u: "https://www.coca-colacompany.com/about-us/history/the-history-of-the-coca-cola-contour-bottle" },
  { t: "Продлённый в 1923 году патент на бутылку датирован 25 декабря, поэтому её прозвали «рождественской».", u: "https://www.coca-colacompany.com/about-us/history/the-history-of-the-coca-cola-contour-bottle" },
  { t: "Первые фигурные бутылки делали из стекла цвета «German Green». Позже этот оттенок переименовали в «Georgia Green».", u: "https://www.coca-colacompany.com/about-us/history/the-history-of-the-coca-cola-contour-bottle" },
  { t: "Исследование 1949 года показало: меньше 1% американцев не могут узнать бутылку кока-колы только по форме.", u: "https://www.coca-colacompany.com/about-us/history/the-history-of-the-coca-cola-contour-bottle" },
  { t: "Знаменитый дизайнер Раймонд Лоуи называл бутылку кока-колы «идеальной упаковкой для жидкости».", u: "https://www.coca-colacompany.com/about-us/history/the-history-of-the-coca-cola-contour-bottle" },
  { t: "Сальвадор Дали изобразил бутылку кока-колы в картине «Поэзия Америки» 1943 года.", u: "https://www.coca-colacompany.com/about-us/history/the-history-of-the-coca-cola-contour-bottle" },
  { t: "Бутылки кока-колы попали в работы Роберта Раушенберга в 1957 году и на выставку Энди Уорхола в 1962-м.", u: "https://www.coca-colacompany.com/about-us/history/the-history-of-the-coca-cola-contour-bottle" },
  { t: "Бутылка 1915 года вмещала 6,5 унции напитка — около 190 мл, а полная весила больше фунта (около 450 г).", u: "https://www.coca-colacompany.com/about-us/history/the-history-of-the-coca-cola-contour-bottle" },
  { t: "Слово «контурная» применительно к бутылке кока-колы впервые появилось во французском журнале в 1925 году.", u: "https://www.coca-colacompany.com/about-us/history/the-history-of-the-coca-cola-contour-bottle" },
  { t: "Кока-кола в банках поступила в продажу в 1960 году.", u: "https://www.coca-cola.com/ch/fr/about-us/1960" },
  { t: "В 1960 году Coca-Cola купила компанию Minute Maid и начала выпускать соки.", u: "https://www.coca-cola.com/ch/fr/about-us/1960" },
  { t: "Sprite появился 1 февраля 1961 года.", u: "https://www.coca-cola.com/ch/fr/about-us/1960" },
  { t: "Первый диетический напиток компании назывался не Diet Coke, а TaB. Он вышел в 1963 году.", u: "https://www.coca-cola.com/ch/fr/about-us/1960" },
  { t: "Diet Coke представили 8 июля 1982 года. Это первый новый бренд со времён 1886 года, получивший товарный знак Coca-Cola. В других странах его называют Coca-Cola Light.", u: "https://www.wikiwand.com/en/Diet_Coke" },
  { t: "В 1965 году Coca-Cola спонсировала первый мультфильм про Чарли Брауна и Снупи — «Рождество Чарли Брауна».", u: "https://www.coca-cola.com/ch/fr/about-us/1960" },
  { t: "Джингл рекламы «Things Go Better with Coke» 1963 года пели The Supremes, Рэй Чарльз, Арета Франклин и Рой Орбисон.", u: "https://www.coca-cola.com/ch/fr/about-us/1960" },
  { t: "В комедии Билли Уайлдера «Раз, два, три» 1961 года Джеймс Кэгни играет руководителя Coca-Cola в Западном Берлине.", u: "https://www.coca-cola.com/ch/fr/about-us/1960" },
  { t: "В 1950 году кока-кола стала первым товаром на обложке журнала Time. Глава компании Роберт Вудрафф отказался от своего портрета: лицо бренда — сам напиток.", u: "https://www.coca-cola.com/ch/fr/about-us/1950" },
  { t: "К 1959 году кока-колу разливали 1700 заводов в более чем 100 странах.", u: "https://www.coca-cola.com/ch/fr/about-us/1950" },
  { t: "На Всемирной выставке 1958 года в Брюсселе у Coca-Cola был павильон с настоящей работающей линией розлива.", u: "https://www.coca-cola.com/ch/fr/about-us/1950" },
  { t: "В 1941 году Роберт Вудрафф распорядился, чтобы любой военный в форме мог купить бутылку кока-колы за 5 центов, где бы он ни служил.", u: "https://www.coca-colacompany.com/about-us/history/the-chronicle-of-coca-cola-a-symbol-of-friendship" },
  { t: "Во время Второй мировой войны вслед за армией за океан отправили 64 разливочных завода. Военные выпили больше 5 миллиардов бутылок кока-колы.", u: "https://www.coca-colacompany.com/about-us/history/the-chronicle-of-coca-cola-a-symbol-of-friendship" },
  { t: "К началу Второй мировой войны кока-колу разливали в 44 странах — по обе стороны будущего фронта.", u: "https://www.coca-colacompany.com/about-us/history/the-chronicle-of-coca-cola-a-symbol-of-friendship" },
  { t: "Fanta появилась в Германии во время Второй мировой войны, когда сироп кока-колы нельзя было ввезти. Первую версию делали из молочной сыворотки и яблочного жмыха.", u: "https://www.heritagedaily.com/2024/08/the-haunting-history-of-fanta/153178" },
  { t: "Название Fanta — от немецкого Fantasie («фантазия»): сотрудников попросили дать волю фантазии, и продавец Джо Книпп предложил «Fanta!».", u: "https://www.heritagedaily.com/2024/08/the-haunting-history-of-fanta/153178" },
  { t: "Сегодняшняя апельсиновая Fanta родом из Неаполя 1950-х. Это первый после кока-колы напиток компании, который стал продаваться по всему миру.", u: "https://www.coca-cola.com/ch/fr/about-us/1950" },
  { t: "Маршал Жуков попросил сделать для него бесцветную кока-колу, похожую на водку. В 1946 году ему доставили 50 ящиков «белой колы» с красной звездой на крышках.", u: "https://www.todayifoundout.com/?p=58980" },
  { t: "Санту для рекламы Coca-Cola с 1931 по 1964 год рисовал Хэддон Сандблом. Он вдохновлялся стихотворением Клемента Мура «Визит Святого Николая» 1823 года.", u: "https://www.coca-colacompany.com/about-us/history/haddon-sundblom-and-the-coca-cola-santas" },
  { t: "Сначала Сандблому позировал друг, отставной продавец Лу Прентисс. Потом художник рисовал Санту с себя, глядя в зеркало.", u: "https://www.coca-colacompany.com/about-us/history/haddon-sundblom-and-the-coca-cola-santas" },
  { t: "Картины Хэддона Сандблома с Сантой выставлялись в Лувре и в Королевском музее Онтарио.", u: "https://www.coca-colacompany.com/about-us/history/haddon-sundblom-and-the-coca-cola-santas" },
  { t: "В 1942 году в рекламе Coca-Cola появился эльф Sprite Boy. Напиток Sprite появился почти через двадцать лет.", u: "https://www.coca-colacompany.com/about-us/history/haddon-sundblom-and-the-coca-cola-santas" },
  { t: "Новогодние грузовики Coca-Cola впервые проехали по экранам зимой 1995 года. Спецэффекты делала Industrial Light & Magic — студия «Звёздных войн».", u: "https://coca-cola.com/ie/en/media-center/the-tale-of-the-christmas-trucks" },
  { t: "Белые медведи впервые появились в рекламе Coca-Cola во Франции в 1922 году, а анимированными стали в ролике «Northern Lights» 1993 года.", u: "https://www.coca-colacompany.com/about-us/history/coca-colas-polar-bears" },
  { t: "Автор анимированных белых медведей Кен Стюарт придумал их, глядя на своего щенка лабрадора, похожего на медвежонка. Голоса медведей — тоже его голос, изменённый компьютером.", u: "https://www.coca-colacompany.com/about-us/history/coca-colas-polar-bears" },
  { t: "Рецепт кока-колы с 1925 года хранился в банковском сейфе в Атланте. Только в 2011 году его перевезли в хранилище музея World of Coca-Cola.", u: "https://investors.coca-colacompany.com/news-events/press-releases/detail/299/coca-cola-moves-its-secret-formula-to-the-world-of-coca-cola" },
  { t: "23 апреля 1985 года Coca-Cola сменила рецепт. После тысяч жалоб — до 8000 звонков в день — уже 11 июля, через 79 дней, вернулась прежняя кола под названием «Coca-Cola Classic».", u: "https://www.history.com/this-day-in-history/april-23/new-coke-debuts-one-of-the-biggest-product-flops-in-history" },
  { t: "В 1985 году кока-колу пили в космосе — на шаттле «Челленджер». Специальная банка обошлась компании примерно в четверть миллиона долларов.", u: "https://www.space.com/space-cola-wars-35-years.html" },
  { t: "Главной проблемой газировки в невесомости астронавты назвали «мокрую отрыжку».", u: "https://www.space.com/space-cola-wars-35-years.html" },
  { t: "Coca-Cola связана с Олимпийскими играми с 1928 года: тогда предприимчивый разливщик привёз в Амстердам 1000 ящиков. Компания — самый давний непрерывный партнёр Игр.", u: "https://olympics.com/ioc/news/coca-cola-and-the-olympic-games-celebrate-90-years-of-partnership" },
  { t: "Coca-Cola поддержала 11 эстафет олимпийского огня, начиная с Барселоны 1992 года, и помогла найти больше 16 000 факелоносцев.", u: "https://olympics.com/ioc/news/coca-cola-and-the-olympic-games-celebrate-90-years-of-partnership" },
  { t: "Бутылки с именами («Share a Coke») впервые появились в Австралии: 150 самых популярных имён. Потом акция прошла больше чем в 70 странах.", u: "https://www.coca-colacompany.com/about-us/history/how-a-campaign-got-its-start-down-under" },
  { t: "В первое лето акции с именами в Австралии продали больше 250 миллионов бутылок и банок — при населении страны меньше 23 миллионов человек.", u: "https://www.coca-colacompany.com/about-us/history/how-a-campaign-got-its-start-down-under" },
  { t: "Знаменитый ролик 1971 года «I'd Like to Buy the World a Coke» начался с нескольких слов, которые креативный директор Билл Бэкер записал на салфетке. Компания получила о нём больше 100 000 писем.", u: "https://www.coca-colacompany.com/about-us/history/id-like-to-buy-the-world-a-coke" },
  { t: "В 1979 году в Пекин пришла первая партия кока-колы — 3000 ящиков. Многие китайцы до этого видели напиток только в американских фильмах.", u: "https://www.whatsonweibo.com/coca-cola-in-china/" },
  { t: "Кокаин убрали из кока-колы вскоре после начала XX века. Листья коки используют до сих пор, но перед этим из них удаляют кокаин.", u: "https://money.yahoo.com/fact-check-cocaine-coke-soda-173900881.html" },
];

/* The same fact for everyone on a given day; a fixed shuffle so neighbouring days are not in date order. */
const FactOfDay = (() => {
  const order = FACTS.map((_, i) => i);
  let seed = 1886;
  const rnd = () => ((seed = (seed * 1103515245 + 12345) % 2147483648) / 2147483648);
  for (let i = order.length - 1; i > 0; i--) { const j = Math.floor(rnd() * (i + 1)); [order[i], order[j]] = [order[j], order[i]]; }
  const dayNumber = (d) => Math.floor(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / 86400000);
  return {
    today(d = new Date()) { return FACTS[order[dayNumber(d) % order.length]]; },
    key(d = new Date()) { return String(dayNumber(d)); },
    count: FACTS.length,
  };
})();
