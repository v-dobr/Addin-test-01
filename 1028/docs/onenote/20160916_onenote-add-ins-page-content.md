# 使用 OneNote 頁面內容 

在 OneNote 增益集 JavaScript API 中，頁面內容會以下列物件模型顯示。

  ![OneNote 頁面物件模型圖](../../images/OneNoteOM-page.png)

- Page 物件包含 PageContent 物件的集合。
- PageContent 物件包含 Outline、Image 或 Other 的內容類型。
- Outline 物件包含 Paragraph 物件的集合。
- Paragraph 物件包含 RichText、Image、Table 或 Other 的內容類型。

若要建立空白的 OneNote 頁面，請使用下列其中一種方法：

- [Section.addPage](../../reference/onenote/section.md#addpagetitle-string)
- [Page.insertPageAsSibling](../../reference/onenote/page.md#insertpageassiblinglocation-string-title-string)

然後在下列物件中使用方法以使用頁面內容，例如 Page.addOutline 和 Outline.appendHtml。 

- [頁面](../../reference/onenote/page.md)
- [大綱](../../reference/onenote/outline.md)
- [段落](../../reference/onenote/paragraph.md)

OneNote 頁面的內容和結構會以 HTML 顯示。 僅支援 HTML 的子集來建立或更新網頁內容，如下所述。

## 支援的 HTML

OneNote 增益集 JavaScript API 支援下列的 HTML 建立和更新網頁內容︰

- `<html>`, `<body>`, `<div>`, `<span>`, `<br/>` 
- `<p>`
- `<img>`
- `<a>`
- `<ul>`, `<ol>`, `<li>` 
- `<table>`, `<tr>`, `<td>`
- `<h1>` ... `<h6>`
- `<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`

## 存取頁面內容

您只能夠透過 `Page#load` 針對目前使用中的頁面存取*頁面內容*。 若要變更使用中的頁面，請叫用 `navigateToPage($page)`。

中繼資料，例如標題，仍然可以針對任何頁面查詢。

## 其他資源

- [OneNote JavaScript API 程式設計的概觀](onenote-add-ins-programming-overview.md)
- [OneNote JavaScript API 參考](../../reference/onenote/onenote-add-ins-javascript-reference.md)
- [Rubric Grader 範例](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office 增益集平台概觀](https://dev.office.com/docs/add-ins/overview/office-add-ins)
