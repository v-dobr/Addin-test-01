# Создание первой надстройки OneNote

В этой статье рассказано, как создать простую надстройку области задач, добавляющую текст на страницу в OneNote.

На приведенном ниже рисунке показана надстройка, которую вы создадите.

   ![Надстройка OneNote, созданная на основе данного пошагового руководства](../../images/onenote-first-add-in.png)

<a name="setup"></a>
## Шаг 1. Настройка среды разработки
1. Установите генератор Yeoman Office и необходимые для него компоненты, выполнив эти [инструкции по установке](https://dev.office.com/docs/add-ins/get-started/create-an-office-add-in-using-any-editor).

   Генератор Yeoman Office упрощает создание проектов надстроек, если у вас нет Visual Studio или вы хотите использовать технологии, отличные от простого HTML, CSS и JavaScript. Кроме того, он обеспечивает быстрый доступ к локальному веб-серверу Gulp для тестирования. 

   >При необходимости для создания файлов проекта вы можете [использовать Visual Studio](https://dev.office.com/docs/add-ins/get-started/create-and-debug-office-add-ins-in-visual-studio), но в этом случае в проекте не будет встроенной поддержки сервера Gulp.

<a name="create-project"></a>
## Шаг 2. Создание проекта надстройки 
1. Создайте локальную папку *onenote add-in*.

2. Откройте **командную строку** и перейдите в папку **onenote add-in**. Запустите команду `yo office`, как показано ниже.

```
C:\your-local-path\onenote add-in\> yo office
```
>Эти инструкции включают действия с использованием командной строки Windows, но их можно выполнять и в других оболочках. 

3. Используя указанные ниже параметры, создайте проект.

| Параметр | Значение |
|:------|:------|
| Имя проекта | Надстройка OneNote |
| Корневая папка проекта | Используйте значение, указанное по умолчанию |
| Тип проекта Office | Надстройка области задач |
| Поддерживаемые приложения Office | Выберите любое приложение, так как мы добавим ведущее приложение OneNote позже |
| Используемые технологии | HTML, CSS и JavaScript |

<a name="manifest"></a>
## Шаг 3. Настройка манифеста надстройки 
1. В папке файлов проекта откройте файл **manifest-onenote-add-in.xml**. Добавьте указанную ниже строку в раздел **Ведущие приложения**. Эта строка указывает, что надстройка поддерживает ведущее приложение OneNote.

```
<Host Name="Notebook" />
```

Обратите внимание, что параметр **SourceLocation** уже настроен для использования веб-сервера Gulp.

```
<SourceLocation DefaultValue="https://localhost:8443/app/home/home.html"/>
```

<a name="develop"></a>
## Шаг 4. Разработка надстройки
Для разработки надстройки вы можете использовать любой текстовый редактор или IDE. Если вы еще не используете Visual Studio Code, вы можете [бесплатно скачать его](https://code.visualstudio.com/) для ОС Linux, Mac OSX и Windows.

1. Откройте файл **home.html** в папке *app/home*. 

2. Измените ссылки на API JavaScript для Office и на стили и компоненты [Office UI Fabric](http://dev.office.com/fabric):

   а) раскомментируйте ссылку fabric.components.min.css;

   б) замените ссылку сценария на Office.js указанной ниже ссылкой на *бета*версию.

```
<script src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
```

Ссылки на Office должны выглядеть указанным ниже образом.

```
<link href="//appsforoffice.microsoft.com/fabric/1.0/fabric.min.css" rel="stylesheet">
<link href="//appsforoffice.microsoft.com/fabric/1.0/fabric.components.min.css" rel="stylesheet">
<script src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
```

3. Замените элемент `<body>` приведенным ниже кодом. При этом с помощью [компонентов Office UI Fabric](http://dev.office.com/fabric/components) будут добавлены текстовая область и кнопка. Структура **Responsive Grid** — из набора [стилей Office UI Fabric](http://dev.office.com/fabric/styles). 

```
<body class="ms-font-m">
    <div class="home flex-container">
        <div class="ms-Grid">
            <div class="ms-Grid-row ms-bgColor-themeDarker">
                <div class="ms-Grid-col">
                    <span class="ms-font-xl ms-fontColor-themeLighter ms-fontWeight-semibold">OneNote Add-in</span>
                </div>
            </div>
        </div>
        <br />
        <div class="ms-Grid">
            <div class="ms-Grid-row">
                <div class="ms-Grid-col">
                    <label class="ms-Label">Enter content here</label>
                    <div class="ms-TextField ms-TextField--placeholder">
                        <textarea id="textBox" rows="5"></textarea>
                    </div>
                </div>
            </div>
            <div class="ms-Grid-row">
                <div class="ms-Grid-col">
                    <div class="ms-font-m ms-fontColor-themeLight header--text">
                        <button class="ms-Button ms-Button--primary" id="addOutline">
                            <span class="ms-Button-icon"><i class="ms-Icon"></i></span>
                            <span class="ms-Button-label">Add outline</span>
                            <span class="ms-Button-description">Adds the content above to the current page.</span>
                        </button>
                    </div>
                </div>
            </div>
        </div>
    </div>
</body>
```

4. Откройте файл **home.js** в папке *app/home*. Измените функцию **Office.initialize** указанным ниже образом, чтобы добавить событие нажатия кнопки **Add outline** (Добавить структуру). 

```
// The initialize function is run each time the page is loaded.
Office.initialize = function (reason) {
    $(document).ready(function () {
        app.initialize();

        // Set up event handler for the UI.
        $('#addOutline').click(addOutlineToPage);
    });
};
```
 
5. Замените метод **getDataFromSelection** указанным ниже методом **addOutlineToPage**. Этот метод получает содержимое из области текста и добавляет его на страницу.

```
// Add the contents of the text area to the page.
function addOutlineToPage() {        
    OneNote.run(function (context) {
       var html = '<p>' + $('#textBox').val() + '</p>';

        // Get the current page.
        var page = context.application.getActivePage();

        // Queue a command to load the page with the title property.             
        page.load('title'); 

        // Add an outline with the specified HTML to the page.
        var outline = page.addOutline(40, 90, html);

        // Run the queued commands, and return a promise to indicate task completion.
        return context.sync()
            .then(function() {
                console.log('Added outline to page ' + page.title);
            })
            .catch(function(error) {
                app.showNotification("Error: " + error); 
                console.log("Error: " + error); 
                if (error instanceof OfficeExtension.Error) { 
                    console.log("Debug info: " + JSON.stringify(error.debugInfo)); 
                } 
            }); 
        });
}
```

<a name="test"></a>
## Шаг 5. Проверка надстройки в OneNote Online
1. Запустите веб-сервер Gulp:  

   a. Откройте командную строку **cmd** в папке **onenote add-in**. 

   b. Запустите команду `gulp serve-static`, как показано ниже.

```
C:\your-local-path\onenote add-in\> gulp serve-static
```

2. Установите самозаверяющий сертификат веб-сервера Gulp в качестве доверенного сертификата. Вам потребуется только один раз сделать это на компьютере, и вы сможете работать с проектами надстроек, созданных с помощью генератора Yeoman Office:

   а) перейдите на страницу размещенной надстройки. По умолчанию это тот же URL-адрес, который используется в манифесте:

```
https://localhost:8443/app/home/home.html
```

   b. установите сертификат в качестве доверенного сертификата. Дополнительные сведения см. в статье [Добавление самозаверяющих сертификатов в качестве доверенных корневых сертификатов](https://github.com/OfficeDev/generator-office/blob/master/docs/trust-self-signed-cert.md).

3. Откройте записную книжку в OneNote Online.

4. Выберите элементы **Вставка > Надстройки Office**. Откроется диалоговое окно "Надстройки Office".
  - Если вы выполнили вход с помощью пользовательской учетной записи, на вкладке **Мои надстройки** выберите элемент **Отправить надстройку**.
  - Если вы выполнили вход с помощью рабочей или учебной учетной записи, на вкладке **Моя организация** выберите элемент **Отправить надстройку**. 
  
  На приведенном ниже изображении показана вкладка **Мои надстройки** для записных книжек отдельного пользователя.

  ![Диалоговое окно "Надстройки Office" со вкладкой "Мои надстройки"](../../images/onenote-office-add-ins-dialog.png)
  
  >**ПРИМЕЧАНИЕ**. Чтобы активировать кнопку **Надстройки Office**, щелкните страницу OneNote.

5. В диалоговом окне "Отправить надстройку" выберите файл **manifest-onenote-add-in.xml** и нажмите кнопку **Отправить**. При тестировании вы можете сохранить файл манифеста в локальном расположении.

6. Надстройка откроется в iFrame рядом со страницей OneNote. Введите текст в текстовой области и нажмите кнопку **Add outline** (Добавить структуру). Введенный текст будет добавлен на страницу. 

## Устранение неполадок и советы
- Для отладки надстройки можно использовать имеющиеся в браузере средства разработчика. При использовании веб-сервера Gulp и отладке в Internet Explorer или Chrome вы можете сохранить внесенные изменения в локальном расположении, а затем просто обновить iFrame надстройки.

- При проверке объекта OneNote для доступных свойств отображаются действительные значения. Для свойств, которые необходимо загрузить, отображается текст *не определено*. Разверните узел `_proto_`, чтобы отобразить свойства, которые определены для объекта, но еще не загружены.

      ![Unloaded OneNote object in the debugger](../../images/onenote-debug.png)

- Если надстройка использует какие-либо HTTP-ресурсы, то вам потребуется включить смешанное содержимое в браузере. Надстройки, которые применяются в рабочей среде, должны использовать только безопасные HTTPS-ресурсы.

## Дополнительные ресурсы

- [Обзор создания кода с помощью API JavaScript для OneNote](onenote-add-ins-programming-overview.md)
- [Справочник по API JavaScript для OneNote](../../reference/onenote/onenote-add-ins-javascript-reference.md)
- [Пример надстройки Rubric Grader](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Обзор платформы надстроек Office](https://dev.office.com/docs/add-ins/overview/office-add-ins)
