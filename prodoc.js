const puppeteer = require("puppeteer");
const ExcelJS = require("exceljs");

const urls = [
  "https://prodoctorov.ru/ulyanovsk/vrach/140776-semenova/",
  "https://prodoctorov.ru/ulyanovsk/vrach/916475-sazonova/",
  "https://prodoctorov.ru/ulyanovsk/vrach/255755-kuznecov/",
  "https://prodoctorov.ru/ulyanovsk/vrach/551348-anfimov/",
  "https://prodoctorov.ru/ulyanovsk/vrach/948028-mayorova/",
  "https://prodoctorov.ru/ulyanovsk/vrach/938035-sveshnikova/",
  "https://prodoctorov.ru/ulyanovsk/vrach/498448-lisyutina/",
  "https://prodoctorov.ru/ulyanovsk/vrach/482730-dombrovskaya/",
  "https://prodoctorov.ru/ulyanovsk/vrach/379426-deryabina/",
  "https://prodoctorov.ru/ulyanovsk/vrach/988756-gricenko/",
  "https://prodoctorov.ru/irkutsk/vrach/263926-kornienko/",
  "https://prodoctorov.ru/irkutsk/vrach/383124-sobennikov/",
  "https://prodoctorov.ru/irkutsk/vrach/653722-merinova/",
  "https://prodoctorov.ru/irkutsk/vrach/725125-petrov/",
  "https://prodoctorov.ru/irkutsk/vrach/862083-skibo/",
  "https://prodoctorov.ru/irkutsk/vrach/101729-sobennikova/",
  "https://prodoctorov.ru/irkutsk/vrach/101728-prokopeva/",
  "https://prodoctorov.ru/irkutsk/vrach/264668-puhlyakova/",
  "https://prodoctorov.ru/irkutsk/vrach/1053495-tihonova/",
  "https://prodoctorov.ru/irkutsk/vrach/101730-chernyak/",
  "https://prodoctorov.ru/chita/vrach/993795-petrov/",
  "https://prodoctorov.ru/chita/vrach/119956-fahrutdinova/",
  "https://prodoctorov.ru/chita/vrach/119188-murzina/",
  "https://prodoctorov.ru/chita/vrach/119961-shilnikov/",
  "https://prodoctorov.ru/chita/vrach/609199-goncharova/",
  "https://prodoctorov.ru/chita/vrach/120210-bodagova/",
  "https://prodoctorov.ru/chita/vrach/338791-davydov/",
  "https://prodoctorov.ru/chita/vrach/814267-kuzmina/",
  "https://prodoctorov.ru/chita/vrach/949426-yuferova/",
  "https://prodoctorov.ru/chita/vrach/839065-titova/",
  "https://prodoctorov.ru/vladivostok/vrach/261134-kovalev/",
  "https://prodoctorov.ru/vladivostok/vrach/261121-belov/",
  "https://prodoctorov.ru/vladivostok/vrach/261136-kolesnik/",
  "https://prodoctorov.ru/vladivostok/vrach/965839-belozerova/",
  "https://prodoctorov.ru/vladivostok/vrach/261135-kovalev/",
  "https://prodoctorov.ru/vladivostok/vrach/487996-bubenko/",
  "https://prodoctorov.ru/vladivostok/vrach/261133-klimenkov/",
  "https://prodoctorov.ru/vladivostok/vrach/259578-gavrilenko/",
  "https://prodoctorov.ru/vladivostok/vrach/1020926-tarasenko/",
  "https://prodoctorov.ru/vladivostok/vrach/159727-sorokin/",
  "https://prodoctorov.ru/vladimir/vrach/327087-zabotin/",
  "https://prodoctorov.ru/vladimir/vrach/710931-neploh/",
  "https://prodoctorov.ru/vladimir/vrach/327031-gerasimova/",
  "https://prodoctorov.ru/vladimir/vrach/92878-aksenova/",
  "https://prodoctorov.ru/vladimir/vrach/565289-ivanova/",
  "https://prodoctorov.ru/vladimir/vrach/92829-alborova/",
  "https://prodoctorov.ru/vladimir/vrach/1089251-shevchuk/",
  "https://prodoctorov.ru/vladimir/vrach/136618-zhutikov/",
  "https://prodoctorov.ru/vladimir/vrach/92717-nikolaev/",
  "https://prodoctorov.ru/vladimir/vrach/916536-nosych/",
  "https://prodoctorov.ru/surgut/vrach/581783-golubkova/",
  "https://prodoctorov.ru/surgut/vrach/581781-ivanov/",
  "https://prodoctorov.ru/surgut/vrach/333097-nafikova/",
  "https://prodoctorov.ru/surgut/vrach/949897-dagaeva/",
  "https://prodoctorov.ru/surgut/vrach/332174-balan/",
  "https://prodoctorov.ru/surgut/vrach/331477-ahverdieva/",
  "https://prodoctorov.ru/surgut/vrach/331804-ignatovich/",
  "https://prodoctorov.ru/surgut/vrach/581784-tuganova/",
  "https://prodoctorov.ru/surgut/vrach/333111-tuganov/",
  "https://prodoctorov.ru/surgut/vrach/333084-buldakova/",
  "https://prodoctorov.ru/tambov/vrach/622730-tolstov/",
  "https://prodoctorov.ru/tambov/vrach/373661-fedotova/",
  "https://prodoctorov.ru/tambov/vrach/952507-volkov/",
  "https://prodoctorov.ru/tambov/vrach/373906-volkov/",
  "https://prodoctorov.ru/tambov/vrach/138012-osnachev/",
  "https://prodoctorov.ru/tambov/vrach/622694-ostrikov/",
  "https://prodoctorov.ru/tambov/vrach/356660-volkova/",
  "https://prodoctorov.ru/tambov/vrach/373660-ogloblin/",
  "https://prodoctorov.ru/tambov/vrach/622751-duhovnikova/",
  "https://prodoctorov.ru/tambov/vrach/622755-egorov/",
  "https://prodoctorov.ru/tula/vrach/1034739-seleznev/",
  "https://prodoctorov.ru/tula/vrach/475462-chernyavskiy/",
  "https://prodoctorov.ru/tula/vrach/1043497-nekrasov/",
  "https://prodoctorov.ru/tula/vrach/792863-borisenko/",
  "https://prodoctorov.ru/tula/vrach/294020-ivashinenko/",
  "https://prodoctorov.ru/tula/vrach/137112-belynceva/",
  "https://prodoctorov.ru/tula/vrach/475434-mihaylovskiy/",
  "https://prodoctorov.ru/tula/vrach/475408-burdelova/",
  "https://prodoctorov.ru/tula/vrach/475152-butuzov/",
  "https://prodoctorov.ru/tula/vrach/294006-shurupov/",
  "https://prodoctorov.ru/habarovsk/vrach/252198-gromova/",
  "https://prodoctorov.ru/habarovsk/vrach/531170-vasilenko/",
  "https://prodoctorov.ru/habarovsk/vrach/124567-kuznecov/",
  "https://prodoctorov.ru/habarovsk/vrach/251445-skripnichenko/",
  "https://prodoctorov.ru/habarovsk/vrach/351807-kravcov/",
  "https://prodoctorov.ru/habarovsk/vrach/531165-tarasov/",
  "https://prodoctorov.ru/habarovsk/vrach/252340-selezneva/",
  "https://prodoctorov.ru/habarovsk/vrach/918074-ratner/",
  "https://prodoctorov.ru/habarovsk/vrach/282812-ryzhkov/",
  "https://prodoctorov.ru/habarovsk/vrach/265139-galynina/",
  "https://prodoctorov.ru/kurgan/vrach/817571-chesnokova/#otzivi",
  "https://prodoctorov.ru/kurgan/vrach/341717-savelev/",
  "https://prodoctorov.ru/kurgan/vrach/109378-pustynnikova/",
  "https://prodoctorov.ru/kurgan/vrach/341702-kuznecova/",
  "https://prodoctorov.ru/kurgan/vrach/819002-sheglov/",
  "https://prodoctorov.ru/kurgan/vrach/341736-kvashnin/",
  "https://prodoctorov.ru/kurgan/vrach/341730-yakupov/",
  "https://prodoctorov.ru/kurgan/vrach/341690-zdorovenko/",
  "https://prodoctorov.ru/kurgan/vrach/341854-telpov/",
  "https://prodoctorov.ru/kurgan/vrach/341355-pankratova/",
  "https://prodoctorov.ru/lipeck/vrach/560604-borovskih/",
  "https://prodoctorov.ru/lipeck/vrach/110673-yakovlev/",
  "https://prodoctorov.ru/lipeck/vrach/121746-melyakov/",
  "https://prodoctorov.ru/lipeck/vrach/668632-grechaninova/",
  "https://prodoctorov.ru/lipeck/vrach/1017451-polikarpov/",
  "https://prodoctorov.ru/lipeck/vrach/570431-gamova/",
  "https://prodoctorov.ru/lipeck/vrach/943406-komarova/",
  "https://prodoctorov.ru/lipeck/vrach/959243-bogomolov/",
  "https://prodoctorov.ru/lipeck/vrach/762497-ivanov/",
  "https://prodoctorov.ru/lipeck/vrach/1041098-bolshenko/",
  "https://prodoctorov.ru/petrozavodsk/vrach/890752-potehina/",
  "https://prodoctorov.ru/petrozavodsk/vrach/363755-kravcova/",
  "https://prodoctorov.ru/petrozavodsk/vrach/565543-lazovskiy/",
  "https://prodoctorov.ru/petrozavodsk/vrach/966831-kelkka/",
  "https://prodoctorov.ru/petrozavodsk/vrach/363745-leshyova/",
  "https://prodoctorov.ru/petrozavodsk/vrach/363762-dmitriev/",
  "https://prodoctorov.ru/petrozavodsk/vrach/816269-abramova/",
  "https://prodoctorov.ru/petrozavodsk/vrach/816224-fofanov/",
  "https://prodoctorov.ru/petrozavodsk/vrach/172337-kargopolceva/",
  "https://prodoctorov.ru/petrozavodsk/vrach/816425-vorobeva/",
  "https://prodoctorov.ru/kursk/vrach/564968-plemenov/",
  "https://prodoctorov.ru/kursk/vrach/301969-kuc/",
  "https://prodoctorov.ru/kursk/vrach/186456-mironyuk/",
  "https://prodoctorov.ru/kursk/vrach/303126-goncharov/",
  "https://prodoctorov.ru/kursk/vrach/301338-bobrovnikova/",
  "https://prodoctorov.ru/kursk/vrach/559633-koneva/",
  "https://prodoctorov.ru/kursk/vrach/559766-timofeev/",
  "https://prodoctorov.ru/kursk/vrach/109983-petrova/",
  "https://prodoctorov.ru/kursk/vrach/301347-bekas/",
  "https://prodoctorov.ru/kursk/vrach/559734-romashova/",
  "https://prodoctorov.ru/penza/vrach/566048-ampleev/",
  "https://prodoctorov.ru/penza/vrach/710738-shnyayder/",
  "https://prodoctorov.ru/penza/vrach/1056425-filippov/",
  "https://prodoctorov.ru/penza/vrach/566439-gabzalilova/",
  "https://prodoctorov.ru/penza/vrach/45134-malkova/",
  "https://prodoctorov.ru/penza/vrach/792656-derevyankina/",
  "https://prodoctorov.ru/penza/vrach/792645-vedeneev/",
  "https://prodoctorov.ru/penza/vrach/569173-atrohov/",
  "https://prodoctorov.ru/penza/vrach/478234-yurkova/",
  "https://prodoctorov.ru/penza/vrach/569313-samarceva/",
  "https://prodoctorov.ru/bryansk/vrach/807408-kozlov/",
  "https://prodoctorov.ru/bryansk/vrach/334745-nikishov/",
  "https://prodoctorov.ru/bryansk/vrach/622664-bibik/",
  "https://prodoctorov.ru/bryansk/vrach/551131-kushev/",
  "https://prodoctorov.ru/bryansk/vrach/603445-bychkova/",
  "https://prodoctorov.ru/bryansk/vrach/914634-savina/",
  "https://prodoctorov.ru/bryansk/vrach/576876-shtutina/",
  "https://prodoctorov.ru/bryansk/vrach/551063-selezneva/",
  "https://prodoctorov.ru/bryansk/vrach/497750-lyubimova/",
  "https://prodoctorov.ru/bryansk/vrach/91423-yunikov/",
  "https://prodoctorov.ru/vologda/vrach/828352-shalagin/",
  "https://prodoctorov.ru/vologda/vrach/1029188-litvinenko/",
  "https://prodoctorov.ru/vologda/vrach/339456-popova/",
  "https://prodoctorov.ru/vologda/vrach/1091528-zagumennova/",
  "https://prodoctorov.ru/vologda/vrach/339734-ershova/",
  "https://prodoctorov.ru/vologda/vrach/339681-duldier/",
  "https://prodoctorov.ru/vologda/vrach/339330-andryayceva/",
  "https://prodoctorov.ru/vologda/vrach/339647-koroleva/",
  "https://prodoctorov.ru/vologda/vrach/828039-kirichenko/",
  "https://prodoctorov.ru/vologda/vrach/339687-saykina/",
  "https://prodoctorov.ru/kaliningrad/vrach/800052-trofimova/",
  "https://prodoctorov.ru/kaliningrad/vrach/194428-ostapchuk/",
  "https://prodoctorov.ru/kaliningrad/vrach/810257-makarova/",
  "https://prodoctorov.ru/kaliningrad/vrach/550214-klimenko/",
  "https://prodoctorov.ru/kaliningrad/vrach/368964-husainov/",
  "https://prodoctorov.ru/kaliningrad/vrach/105492-solovev/",
  "https://prodoctorov.ru/kaliningrad/vrach/547783-golopuro/",
  "https://prodoctorov.ru/kaliningrad/vrach/590462-polyanskaya/",
  "https://prodoctorov.ru/kaliningrad/vrach/966051-marinova/",
  "https://prodoctorov.ru/kaliningrad/vrach/917447-lisogurskaya/",
  "https://prodoctorov.ru/oryol/vrach/339953-aldoshin/",
  "https://prodoctorov.ru/oryol/vrach/719836-fishev/",
  "https://prodoctorov.ru/oryol/vrach/340384-knyazeva/",
  "https://prodoctorov.ru/oryol/vrach/153220-samonova/",
  "https://prodoctorov.ru/oryol/vrach/340375-karpuhin/",
  "https://prodoctorov.ru/oryol/vrach/993689-pylina/",
  "https://prodoctorov.ru/oryol/vrach/815833-smirnov/",
  "https://prodoctorov.ru/oryol/vrach/54493-koroleva/",
  "https://prodoctorov.ru/oryol/vrach/196640-vasilenko/",
  "https://prodoctorov.ru/oryol/vrach/815814-lanceva/",
  "https://prodoctorov.ru/tomsk/vrach/629546-francuzenko/",
  "https://prodoctorov.ru/tomsk/vrach/388301-frenovskaya/",
  "https://prodoctorov.ru/tomsk/vrach/282660-seryh/",
  "https://prodoctorov.ru/tomsk/vrach/283140-minevich/",
  "https://prodoctorov.ru/tomsk/vrach/282064-mamysheva/",
  "https://prodoctorov.ru/tomsk/vrach/338926-solomatin/",
  "https://prodoctorov.ru/tomsk/vrach/1027761-schastnyy/",
  "https://prodoctorov.ru/tomsk/vrach/282662-yakimova/",
  "https://prodoctorov.ru/tomsk/vrach/388282-mochalova/",
  "https://prodoctorov.ru/tomsk/vrach/580800-bochkareva/",
  "https://prodoctorov.ru/barnaul/vrach/90194-makarov/",
  "https://prodoctorov.ru/barnaul/vrach/536036-grebenshikov/",
  "https://prodoctorov.ru/barnaul/vrach/235136-zayukova/",
  "https://prodoctorov.ru/barnaul/vrach/665666-zateev/",
  "https://prodoctorov.ru/barnaul/vrach/160309-portnyagina/",
  "https://prodoctorov.ru/barnaul/vrach/235140-gubaydulina/",
  "https://prodoctorov.ru/barnaul/vrach/957716-zheleznaya/",
  "https://prodoctorov.ru/barnaul/vrach/587345-goncharova/",
  "https://prodoctorov.ru/barnaul/vrach/536618-deniko/",
  "https://prodoctorov.ru/barnaul/vrach/965049-yuzhaninova/",
  "https://prodoctorov.ru/ryazan/vrach/564531-turbina/",
  "https://prodoctorov.ru/ryazan/vrach/527458-shablovskaya/",
  "https://prodoctorov.ru/ryazan/vrach/137763-lukashuk/",
  "https://prodoctorov.ru/ryazan/vrach/128098-shutova/",
  "https://prodoctorov.ru/ryazan/vrach/165321-ageenko/",
  "https://prodoctorov.ru/ryazan/vrach/699868-panov/",
  "https://prodoctorov.ru/ryazan/vrach/137810-novikov/",
  "https://prodoctorov.ru/ryazan/vrach/128168-averyanov/",
  "https://prodoctorov.ru/ryazan/vrach/128204-filimonov/",
  "https://prodoctorov.ru/ryazan/vrach/188253-novikov/",
  "https://prodoctorov.ru/kaluga/vrach/216038-kambulova/",
  "https://prodoctorov.ru/kaluga/vrach/816998-chernikova/",
  "https://prodoctorov.ru/kaluga/vrach/722489-voroncov/",
  "https://prodoctorov.ru/kaluga/vrach/884664-doncova/",
  "https://prodoctorov.ru/kaluga/vrach/831355-vedernova/",
  "https://prodoctorov.ru/kaluga/vrach/337689-radulov/",
  "https://prodoctorov.ru/kaluga/vrach/221383-arshanskiy/",
  "https://prodoctorov.ru/kaluga/vrach/816998-chernikova/",
  "https://prodoctorov.ru/kaluga/vrach/221383-arshanskiy/",
  "https://prodoctorov.ru/kaluga/vrach/106357-bizyulev/",
  "https://prodoctorov.ru/smolensk/vrach/1028148-razina/",
  "https://prodoctorov.ru/smolensk/vrach/712708-hraimenkova/",
  "https://prodoctorov.ru/smolensk/vrach/449999-tedeeva/",
  "https://prodoctorov.ru/smolensk/vrach/336830-filimonova/",
  "https://prodoctorov.ru/smolensk/vrach/392291-bykova/",
  "https://prodoctorov.ru/smolensk/vrach/148641-denchenkov/",
  "https://prodoctorov.ru/smolensk/vrach/140694-bozhenkov/",
  "https://prodoctorov.ru/smolensk/vrach/471729-sidochenkov/",
  "https://prodoctorov.ru/smolensk/vrach/336805-maslyanaya/",
  "https://prodoctorov.ru/smolensk/vrach/816084-kuzmin/",
  "https://prodoctorov.ru/pskov/vrach/353351-sacevich/",
  "https://prodoctorov.ru/pskov/vrach/360365-indyukov/",
  "https://prodoctorov.ru/pskov/vrach/975248-petrova/",
  "https://prodoctorov.ru/pskov/vrach/360363-chechkin/",
  "https://prodoctorov.ru/pskov/vrach/682000-moskalev/",
  "https://prodoctorov.ru/pskov/vrach/360367-petrova/",
  "https://prodoctorov.ru/pskov/vrach/360368-krupenenkova/",
  "https://prodoctorov.ru/pskov/vrach/460689-gavryuchenkov/",
  "https://prodoctorov.ru/pskov/vrach/360099-kukareko/",
  "https://prodoctorov.ru/pskov/vrach/617508-eyberman/",
]; // Замените URL1, URL2, URL3 на ваши ссылки
const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet("DoctorsData");
worksheet.columns = [
  { header: "ФИО врача", key: "name" },
  { header: "Стаж работы", key: "experience" },
  { header: "Фото врача", key: "photo" },
  { header: "Период работы в клинике", key: "period" },
  { header: "Образование", key: "education" },
  { header: "Расписание", key: "schedule" },
  { header: "Должность", key: "position" },
  { header: "Специализация", key: "specialization" },
  { header: "Повышение квалификации", key: "qualification" },
  { header: "Дипломы и сертификаты", key: "diploma" },
  { header: "ProDoctorov", key: "url" },
  { header: "Другие агрегаторы", key: "other" },
];

(async () => {
  const browser = await puppeteer.launch();
  const page = await browser.newPage();

  for (const url of urls) {
    await page.goto(url);
    console.log(`going to ${url}`);
    const name = await page.$eval("h1", (element) => element.textContent);
    const experience = await page.$eval(
      "div.ui-text_subtitle-2",
      (element) => element.textContent
    );
    const photo = await page.$eval('img[itemprop="image"]', (element) =>
      element.getAttribute("src")
    );
    const education = await page.$eval("#educations", (element) => {
      const coursesList = element.querySelectorAll("li");
      const qualificationArray = [];
      coursesList.forEach((course) => {
        qualificationArray.push(course.textContent);
      });
      return qualificationArray.join("\n");
    });
    const specialization = await page.$eval(
      ".b-doctor-intro__specs.mb-4",
      (element) => {
        const coursesList = element.querySelectorAll("a");
        const qualificationArray = [];
        coursesList.forEach((course) => {
          qualificationArray.push(course.textContent);
        });
        return qualificationArray.join(", ");
      }
    );
      let qualification;
    // const specialization = await page.$eval('.b-doctor-intro__specs.mb-4', (element) => element.textContent);
    try {
       qualification = await page.$eval("#courses", (element) => {
        const coursesList = element.querySelectorAll("li");
        const qualificationArray = [];
        coursesList.forEach((course) => {
          qualificationArray.push(course.textContent);
        });
        return qualificationArray.join("\n");
      });
      console.log(qualification); // Вывод квалификаций, если элемент найден
    } catch (error) {
      console.error("Элемент не найден на странице");
      // Можно добавить дополнительные действия, если нужно
    }

    worksheet.addRow({
      name,
      experience,
      photo,
      education,
      specialization,
      qualification,
      url,
    });
  }

  await workbook.xlsx.writeFile("DoctorsData.xlsx");
  await page.waitForTimeout(4000);
  await browser.close();
})();
