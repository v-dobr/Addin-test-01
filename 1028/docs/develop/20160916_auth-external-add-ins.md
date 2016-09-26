# 在您的 Office 增益集授權外部服務

受歡迎的線上服務，包括 Office 365、Google、Facebook、LinkedIn、SalesForce 和 GitHub，讓開發人員給予使用者對於他們在其他應用程式中帳戶的存取權。 這可讓您將這些服務包含在您的 Office 增益集中。 

讓 Web 應用程式存取至線上服務的業界標準架構稱為 OAuth 2.0。 在大部分情況下，您不需要知道架構如何運作的詳細資料，就可以在您的增益集中使用。 許多程式庫均可用，為您摘要詳細資料。

OAuth 的基本概念是應用程式可以是本身的安全性主體，就像是使用者或群組，具有它自己的身分識別與權限集合。 在最常見的案例中，當使用者在需要線上服務的 Office 增益集採取動作時，增益集傳送一組特定權限的服務要求給使用者帳戶。 然後，服務則會提示使用者將這些權限授與增益集。 授與權限之後，服務會將小型編碼*存取權杖*傳送給增益集。 增益集可以使用服務，方法是將它的所有要求的權杖包含至服務的 API。 但是增益集只能在使用者授與的權限內進行操作。 權杖也會在指定的時間之後到期。

數個 OAuth 模式，稱為*流程*或*授與類型*，是專為不同的案例設計。 以下是兩個最重要的模式︰

- **隱含流程**：增益集與線上服務之間的通訊是使用用戶端 JavaScript 來實作。
- **授權程式碼流程**：通訊是增益集的 Web 應用程式與線上服務之間的*伺服器對伺服器*。 因此，它是使用伺服器端程式碼實作。

流程的目的是保護應用程式的身分識別與授權。 在授權程式碼流程中，您會看到必須保持隱藏的*用戶端密碼*。 單一頁面應用程式 (SPA) 無法保護密鑰，所以我們建議您在 SPA 中使用隱含流程。 

您應該熟悉兩個流程的優點及缺點。 [授權程式碼](https://tools.ietf.org/html/rfc6749#section-1.3.1)和[隱含](https://tools.ietf.org/html/rfc6749#section-1.3.2)的正式定義是很好的起點。 

>**附註：**您也可以選擇讓 middleman 服務為您執行所有的授權，並且將存取權杖傳遞至增益集。 如需詳細資訊，請參閱本文稍後的 *Middleman 服務*章節。

## 在 Office 增益集中使用隱含流程
找出線上服務是否支援隱含流程的最佳方式是參考說明文件。

針對支援它的服務，我們會提供 JavaScript 程式庫，為您進行所有詳細工作︰

[Office-js-helpers](https://github.com/OfficeDev/office-js-helpers)

儲存機制的 \demo 資料夾包含範例增益集，使用程式庫以存取一些受歡迎的服務，包括 Google、Facebook 和 Office 365。

另請參閱本文稍後的**程式庫**章節。

## 在 Office 增益集中使用授權程式碼流程

我們有一些範例增益集使用授權程式碼流程︰

- [Office-Add-in-Nodejs-ServerAuth](https://github.com/OfficeDev/Office-Add-in-Nodejs-ServerAuth) (NodeJS)
- [PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) (ASP.NET MVC)

許多程式庫可用於以不同語言和架構來實作授權程式碼流程。 如需詳細資訊，請參閱本文稍後的**程式庫**章節。

### 轉送/Proxy 函式

您甚至可以搭配使用授權程式碼流程與無伺服器 Web 應用程式，方法是在簡單的函式中儲存*用戶端 ID*和*用戶端密碼*值，該函式裝載於如 [Azure Functions](https://azure.microsoft.com/en-us/services/functions) 或 [Amazon Lambda](https://aws.amazon.com/lambda) 的服務。
函式交換適當*存取權杖*的指定程式碼，並將它轉送回用戶端。 這種方法的安全性取決於函式存取權的受保護程度。

若要使用這項技術，您的增益集會顯示 UI/快顯功能表，以顯示線上服務的登入畫面 (例如 Google、Facebook 等等)。 當使用者登入，並授與增益集權限給她在線上服務的資源時，開發人員就會收到可以再傳送給線上函式的程式碼。 本文中的 **Middleman 服務**中說明的服務使用類似這個的流程。 

## 程式庫

程式庫適用於許多語言與平台，並且適用於這兩種流程。 有些是一般用途，其他則是針對特定的線上服務。 

**Office 365 和其他使用 Azure Active Directory 做為授權提供者的服務**：[Azure Active Directory 驗證程式庫](https://azure.microsoft.com/en-us/documentation/articles/active-directory-authentication-libraries/)。 預覽也適用於 [Microsoft 驗證程式庫](https://www.nuget.org/packages/Microsoft.Identity.Client)。

**Google**：針對「驗證」或您的語言的名稱搜尋 [GitHub.com/Google](https://github.com/google)。 大部分的相關儲存機制名為 `google-auth-library-[name of language]`。

**Facebook**：針對「程式庫」或「sdk」搜尋[適用於開發人員的 Facebook](https://developers.facebook.com)。 

**一般 OAuth 2.0**：超過十幾種語言的程式庫的連結頁面是由 IETF OAuth 工作群組在下列位置維護︰[OAuth 程式碼](http://oauth.net/code/)。 請注意，這些程式庫的部分是用來實作 OAuth 相容服務。 身為增益集開發人員的您有興趣的程式庫，在這個頁面上稱為*用戶端*程式庫，因為您的 Web 伺服器是 OAuth 相容服務的用戶端。

## Middleman 服務

您的增益集可以使用 middleman 服務，例如 Auth0，提供許多受歡迎線上服務的存取權杖，或是簡化啟用增益集的社交登入程序，或是兩者。 使用極少的程式碼，增益集可以使用用戶端指令碼或伺服器端程式碼，以連接到 middleman，它會傳送回線上服務的任何必要權杖。 所有的授權實作程式碼位於 middleman 服務。 

我們有一個範例，使用 Auth0 以啟用具有 Facebook、Google 和 Microsoft 帳戶的社交登入︰

[Office-Add-in-Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0)

## 什麼是 CORS？

CORS 代表[跨原始來源資源共用](https://developer.mozilla.org/en-US/docs/Web/HTTP/Access_control_CORS)。 如需有關如何使用增益集內的 CORS 的詳細資訊，請參閱[解決 Office 增益集中的相同原始來源原則的限制](http://dev.office.com/docs/add-ins/develop/addressing-same-origin-policy-limitations)。
