# 在 Office 外接程序中授权外部服务

热门联机服务（包括 Office 365、Google、Facebook、LinkedIn、SalesForce 和 GitHub）使开发人员能够让用户在其他应用程序中访问他们的帐户。 这使你可以在 Office 外接程序中提供这些服务。 

启用 Web 应用程序对联机服务访问的行业标准框架被称为 OAuth 2.0。 在大多数情况下，你不需要了解框架在外接程序中使用原理的详细信息。 许多库都可用来为你提取详细信息。

OAuth 的基本概念是，应用程序本身可以是一个安全主体，就像一个用户或组，拥有其自己的标识和权限集。 在最典型的应用场景中，当用户在需要联机服务的 Office 外接程序中进行操作时，外接程序会向服务发送请求，请求为用户帐户提供一组特定权限。 然后，该服务会提示用户向外接程序授予这些权限。 授予权限之后，该服务会向外接程序发送一个小的编码*访问令牌*。 外接程序可以通过在其向服务 API 发送的所有请求中包含令牌来使用该服务。 但外接程序只能在用户授予它的权限范围内进行操作。 令牌还会在某个指定时间后过期。

称为*流*或*授权类型*的几个 OAuth 模式专为不同应用场景而设计。 以下是两个最重要的模式：

- **隐式流**：外接程序和联机服务之间的通信通过客户端 JavaScript 实现。
- **授权代码流**：外接程序的 Web 应用程序和联机服务之间的通信是*服务器到服务器*。 因此，它是通过服务器端代码实现。

这些流的目的是保护应用程序的标识和授权。 在授权代码流中，提供了一个需要保持隐藏的*客户端密码*。 单页应用程序 (SPA) 无法保护密码，因此建议在 SPA 中使用隐式流。 

你应该熟悉这两个流的其他优缺点。 从[授权代码](https://tools.ietf.org/html/rfc6749#section-1.3.1)和[隐式](https://tools.ietf.org/html/rfc6749#section-1.3.2)中的官方定义开始了解这两个流。 

>**注意：**还可以选择中间人服务为你执行所有授权，并将访问令牌传递给外接程序。 有关详细信息，请参阅本文后面部分中的*中间人服务*。

## 在 Office 外接程序中使用隐式流
了解联机服务是否支持隐式流的最好办法是查阅本文档。

对于支持隐式流的服务，我们提供 JavaScript 库，它将为你完成所有详细工作：

[Office-js-helpers](https://github.com/OfficeDev/office-js-helpers)

存储库的 \demo 文件夹中包含使用库来访问某些热门服务（包括 Google、Facebook 和 Office 365）的示例外接程序。

请参阅本文后面中的**库**部分。

## 在 Office 外接程序中使用授权代码流

我们有一些使用授权代码流的示例外接程序：

- [Office-Add-in-Nodejs-ServerAuth](https://github.com/OfficeDev/Office-Add-in-Nodejs-ServerAuth) (NodeJS)
- [PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) (ASP.NET MVC)

许多库可用于在各种语言和框架中实现授权代码流。 有关详细信息，请参阅本文后面中的**库**部分。

### 中继/代理函数

通过存储托管在如 [Azure 函数](https://azure.microsoft.com/en-us/services/functions)或 [Amazon Lambda](https://aws.amazon.com/lambda) 服务的简单函数中的*客户端 ID* 和*客户端密码*值，甚至可以在无服务器的 Web 应用程序上使用授权代码流。
函数将以给定的代码交换适当的*访问令牌*，并将其重新传递给客户端。 这种方法的安全性取决于对函数访问的保护程度。

若要使用此技术，外接程序会显示 UI/弹出窗口来显示联机服务（如 Google、Facebook 等）的登录屏幕。 当用户登录并对外接程序授予其联机服务中的资源时，开发人员会收到一个代码，然后可以将其发送给联机函数。 本文**中间人服务**中描述的服务使用与此类似的流。 

## 库

库可用于多种语言和多个平台，并且均可用于这两个流。 一些库适用于一般用途，有些库则针对的是特定的联机服务。 

**将 Azure Active Directory 作为授权提供程序使用的 Office 365 及其他服务**：[Azure Active Directory 授权库](https://azure.microsoft.com/en-us/documentation/articles/active-directory-authentication-libraries/)。 预览也适用于 [Microsoft 身份验证库](https://www.nuget.org/packages/Microsoft.Identity.Client)。

**Google**：在 [GitHub.com/Google](https://github.com/google) 中搜索 "auth" 或你语言的相应名称。 大部分的相关存储库被命名为 `google-auth-library-[name of language]`。

**Facebook**：在 [Facebook 开发者](https://developers.facebook.com) 中搜索 "library" 或 "sdk"。 

**常规 OAuth 2.0**：指向十几种语言库的链接页面由 IETF OAuth 工作组在以下位置进行维护：[OAuth 代码](http://oauth.net/code/)。 请注意，其中一些库可用来实现 OAuth 兼容服务。 作为外接程序开发人员，你所感兴趣的库就是此页上称为*客户端*的库，因为 Web 服务器是 OAuth 兼容服务的客户端。

## 中间人服务

外接程序可以使用一个中间人服务（如 Auth0）为许多热门联机服务提供访问令牌或简化启用外接程序社交登录的进程，或同时具备这两个功能。 通过极少量的代码，外接程序可以使用客户端脚本或服务器端代码连接到中间人，并且它将为联机服务发送回任何所需的令牌。 所有授权实现代码都在中间人服务中。 

我们有一个示例，它使用 Auth0 来启用 Facebook、Google 和 Microsoft 帐户的社交登录：

[Office-Add-in-Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0)

## 什么是 CORS？

CORS 代表[跨域资源共享](https://developer.mozilla.org/en-US/docs/Web/HTTP/Access_control_CORS)。 有关如何使用外接程序内的 CORS 进行工作的信息，请参阅[在 Office 外接程序中解决同源策略限制](http://dev.office.com/docs/add-ins/develop/addressing-same-origin-policy-limitations)。
