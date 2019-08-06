<!DOCTYPE html>
<html>
<head>
        <meta charset="UTF-8">
        <title>Editor</title>
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <meta name="description" content="Editor">
        <meta name="theme-color" content="theme">
        <meta content='id' name='language'/>
        <meta content='id' name='geo.country'/>
        <meta content='Indonesia' name='geo.placename'/>
        <meta name="keywords" content="Editor">
        <meta name="author" content="Harjito, harjito@mail.unnes.ac.id">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.3.1/css/bootstrap.min.css" integrity="sha256-YLGeXaapI0/5IgZopewRJcFXomhRMlYYjugPLSyNjTY=" crossorigin="anonymous" />
		<link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.css" rel="stylesheet">
<style>
body{
    padding: 20px 20px 20px 20px;
}
#result{
    border: 1px solid #000000;
    margin: 10px 10px 10px 10px;
    padding: 10px 10px 10px 10px;
}
p{
    margin-left: 80px;
    margin-right:80px;
}
.Title{
    font-size: 24px;
    font-weight:bold;
    text-align:center;
}
.Authors{
    font-size:18px;
    text-align:center;
}
.Addresses{
    font-size:18px;
    text-align:center;
}
.Email{
    font-size:18px;
    text-align:center;
}
.Abstract{
    font-size:18px;
    text-align:justify;
    padding-left:80px;
    padding-right:80px;
}
.Section{
    font-size:22px;
    font-weight:bold;
    list-style-type: upper-alpha;
}
.Bodytext{
    font-size:18px;
    text-align:justify;
}
.BodytextIndented{
    font-size:18px;
    text-align:justify;
    text-indent:80px;
}
.borderTop{
    border-top: 1px solid #1b0101;
}
.borderBottom{
    border-bottom: 1px solid #1b0101;
}
</style>
</head>
<body>
    <h1 class='title'>Convert DOCX to HTML</h1>
    <h3>Choose File</h3>
    <button id="reload" class="button">Reload</button>
    <input type="file" id="file" name="file" multiple /><br>
    <div id="result_block" class="hidden">
        <h3>Content</h3>
        <div id="result">Ready</div>
    </div>
</body>
<script src="/js/jquery.min.js"></script>
<script src="/js/jszip.min.js"></script>
<script src="/js/app.js?v=<?=filemtime('js/app.js')?>"></script>
</html>