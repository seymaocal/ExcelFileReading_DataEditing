<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="index.aspx.cs" Inherits="WebApplication2.WebForm1"  %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" lang="tr">

<head runat="server">
<meta http-equiv="Content-Type" content="text/html" charset="utf-8">

    <title></title>
    
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" />
    <link href="Content/bootstrap.min.css" rel="stylesheet" />
    <script src="Scripts/jquery-3.0.0.min.js"></script>
    <script src="Scripts/bootstrap.min.js"></script>


    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css" />
   <script type="text/javascript" src="http://code.jquery.com/jquery-latest.min.js"></script>


    <!-- datasource !-->
    <link rel="" href="https://cdn.datatables.net/1.10.19/css/jquery.dataTables.min.css" />
    <script src="https://code.jquery.com/jquery-3.3.1.js "></script>
    <script src="https://cdn.datatables.net/1.10.19/js/jquery.dataTables.min.js "></script>
    
    <script>
        $(function () {
            $("#gridColon2").prepend($("<thead></thead>").append($(this).find("tr:first"))).dataTable();
        });
    </script>


  
 

    <style type="text/css">
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
.table table tbody tr td a,
        .table table tbody tr td span {
            position: relative;
            float: left;
            padding: 6px 12px;
            margin-left: -1px;
            line-height: 1.42857143;
            color: #337ab7;
            text-decoration: none;
            background-color: #fff;
            border: 1px solid #ddd;
        }

        .table table > tbody > tr > td > span {
            z-index: 3;
            color: #fff;
            cursor: default;
            background-color: #337ab7;
            border-color: #337ab7;
        }

        .table table > tbody > tr > td :first-child > a,
        .table table > tbody > tr > td:first-child > span {
            margin-left: 0;
            border-top-left-radius: 4px;
            border-bottom-left-radius: 4px;
        }

        .table table > tbody > tr > td:last-child > a,
        .table table > tbody > tr > td:last-child > span {
            border-top-right-radius: 4px;
            border-bottom-right-radius: 4px;
        }

        .table table > tbody > tr > td > a:hover,
        .table table > tbody > tr > td > span:hover,
        .table table > tbody > tr > td > a:focus,
        .table table > tbody > tr > td > span:focus {
            z-index: 2;
            color: #23527c;
            background-color: #eee;
            border-color: #ddd;
        }        

        .btn {
                text-align: center;
                
                font-size: 12px;
            }

        .fileupload {
            font-size: 12px;
        }
         .radiolist {
            font-size: larger;
        }

            .radiolist .btn {
                text-align: center;
                width: 80px;
                height: 25px;
                font-size: 12px;
            }


        * {
            box-sizing: border-box;
            font-size: 12px;
        }


        .fileuploads {
            float: left;
        }

        .btnupload {  
            
            font-size: 13px;
            margin-left:5px;
        }

        .genel {
            margin-top: 10px;
        }

       

        .btn-primary {
            color: #fff;
            background-color: #337ab7;
            border-color: #337ab7;
        }

            .btn-primary:hover {
                color: #fff;
                background-color: #337ab7;
                border-color: #337ab7;
            }

            .btn-primary:not(:disabled):not(.disabled):active, .btn-primary:focus {
                color: #fff;
                background-color: #337ab7;
                border-color: #337ab7;
            }

            
             .alert-secondary {
             color: #383d41;
             background-color: white;
             font-size: 12px;
             font-weight: bold;
             border-color:white;
             margin-top:5px;
             }
.load-wrapp {
    display:none;
    float: left;
    width: 100px;
    height: 100px;
    margin: 0 10px 10px 0;
    padding: 20px 20px 20px;
    border-radius: 5px;
    text-align: center;
}
.load-wrapp:last-child {margin-right: 0;}
.line {
    display: inline-block;
    width: 15px;
    height: 15px;
    border-radius: 15px;
    background-color: #4b9cdb;
}

.load-3 .line:nth-last-child(1) {animation: loadingC .6s .1s linear infinite;}
.load-3 .line:nth-last-child(2) {animation: loadingC .6s .2s linear infinite;}
.load-3 .line:nth-last-child(3) {animation: loadingC .6s .3s linear infinite;}
@keyframes loadingC {
    0 {transform: translate(0,0);}
    50% {transform: translate(0,15px);}
    100% {transform: translate(0,0);}
}

.loading
    {
        font-family: Arial;
        font-size: 10pt;
        width: 200px;
        height: 100px;
        display: none;
        position: fixed;
        background-color:transparent;
        z-index: 999;
    }

.alert{
    background-color:transparent;
    border:none;
    color:red;
    font-size:15px;
    float:center;
  
}
.klavuz{float:right; margin-left:5px;}
.btn_drop{width:fit-content; float:left;}

.FixedHeader {
            position: absolute;
            font-weight: bold;
        }    


.datatables_info{
    margin:5px;
}
a{

    margin:5px;
    cursor:pointer;
}

input[type=search]{
    text-transform:uppercase;
}

</style>

 
</head>

<body>

    <form id="form1" method="post" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server">
        </asp:ScriptManager>
        <div class="container">
        <div class="genel">
            
       
            <asp:FileUpload ID="fluExcel" runat="server" CssClass="fileupload" />


            <asp:Button ID="btnUpload" runat="server" Text="Yükle" OnClick="btnUpload_Click"  CssClass=" btn-submit btn btn-primary btnupload" type="submit" /> 
            <asp:DropDownList ID="dropdownlist" runat="server" Width="200px" Height="28px" Font-Size="12px">
                <asp:ListItem Selected="True" Value="0">Banka Seçiniz</asp:ListItem>
                <asp:ListItem  Value="1">Akbank</asp:ListItem>
                <asp:ListItem  Value="2">Albaraka</asp:ListItem>
                <asp:ListItem  Value="3">Halkbank</asp:ListItem>
                <asp:ListItem  Value="4">Kuveyt</asp:ListItem>
                <asp:ListItem  Value="5">Ptt</asp:ListItem>
                <asp:ListItem  Value="6">Vakýf Katýlým</asp:ListItem>
                <asp:ListItem  Value="7">Vakýf Bank</asp:ListItem>
                <asp:ListItem  Value="8">Ziraat Bank</asp:ListItem>
                <asp:ListItem  Value="9">Ziraat Katýlým</asp:ListItem>
            </asp:DropDownList>
       
            <a href="KullanmaKlavuzu.aspx" class="btn btn-primary btnupload klavuz" target="_blank" >Kullanma Klavuzu</a> 
          
          <asp:Button ID="kaydet" runat="server" Text="Ýndir" OnClick="kaydet_Click" CssClass=" btn btn-primary btnupload" />
            <div class="alert_genel" style="height:50px; background-color:transparent;">
          <div class="alert alert-danger" role="alert">
            <asp:Label ID="lblMessage" runat="server" Text="" ></asp:Label>  </div>  
           </div>      
        </div>
           

  <div class="load-wrapp loading" >
     <div class="load-3">               
       <div class="line"></div>
       <div class="line"></div>
       <div class="line"></div>       
       
     </div>
   </div>         
          
                                
        <div>
            <!--
            <asp:GridView ID="grdExcel" runat="server"></asp:GridView> !-->
        </div>      
            
            
        <div>
            <asp:GridView ID="gridColon2" runat="server" CssClass="table table-striped table-bordered table-hover  ">   
                   
            </asp:GridView> 
           
            
        </div>
</div>     
    </form>

 <script>
$(document).ready(function(){

    setTimeout(function(){

        $("div.alert").fadeOut("slow", function () {

        $("div.alert").remove();

    });

}, 5000);

});
</script>


  <!--Loading !-->

 <script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.8.3/jquery.min.js"></script>


<script type="text/javascript">

    var $x = jQuery.noConflict(); // takma ad oluþturma
    $('.btn-submit').click(function (e) {
        function ShowProgress() {
            setTimeout(function () {
                var modal = $x('<div />');
                modal.addClass("modal");
                $x('body').append(modal);
                var loading = $x(".loading");
                loading.show();
                var top = Math.max($(window).height() / 2 - loading[0].offsetHeight / 2, 0);
                var left = Math.max($(window).width() / 2 - loading[0].offsetWidth / 2, 0);
                loading.css({ top: top, left: left });
            }, 200);
        }


        $x('form').live("submit", function () {
            ShowProgress();
        });

    });
</script>



   



</body>
</html>
