<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="KullanmaKlavuzu.aspx.cs" Inherits="WebApplication2.KullanmaKlavuzu" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" />
    <link href="Content/bootstrap.min.css" rel="stylesheet" />
    <script src="Scripts/jquery-3.0.0.min.js"></script>
    <script src="Scripts/bootstrap.min.js"></script>
    <title></title>
    <style type="text/css">
.container{border: 1px solid dimgray;}
body {
    margin: 0 auto;
    padding: 20px;
    max-width: 1200px;
    overflow-y: scroll;
    font-family: 'Open Sans',sans-serif;
    font-weight: 400;
    color: #777;
    background-color: #f7f7f7;
    -webkit-font-smoothing: antialiased;
    -webkit-text-size-adjust: 100%;
    -ms-text-size-adjust: 100%;
}
.baslik{text-align:center}
p{font-size: 12px;}
li{font-size: 12px;}

    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div class="container">
        
       <h1 class="baslik">Kullanma Klavuzu</h1>
        <p>1. Bu uygulamanın kullanılabilmesi için yüklenen dosyaların excel dosyası olması gerekmektedir.</p>
        <p>2. Excel dosyalarının isimleri büyük harfle başlamalı ve aşağıdaki formatta olması gerekmektedir.
            <ul>
            <li>Akbank</li>
            <li>Albaraka</li>
            <li>Halkbank</li>
            <li>Kuveyt</li>
            <li>Ptt</li>
            <li>Vakıf Bank</li>
            <li>Vakıf Katılım</li>
            <li>Ziraat Bank</li>
            <li>Ziraat Katılım</li>
            </ul>
        </p>
        <p>3. Düzenlenmesi istenilen bilgilerin excel dosyasının, B kolonunda, Tutar bilgisinin ise C kolonunda bulunması gerekmektedir.</p>
        <p>4. Düzenlenen excel dosyasına indir butonuna basıp, bilgisayarınıza indirerek ulaşabilirsiniz.</p>
    </div>
        
    </form>
</body>
</html>
