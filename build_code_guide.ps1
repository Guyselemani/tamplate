$ErrorActionPreference = 'Stop'

$outputPath = Join-Path $PSScriptRoot 'Explication_codes_HTML_CSS.docx'
$tempRoot = Join-Path $PSScriptRoot '.docx-build'
if (Test-Path -LiteralPath $tempRoot) { Remove-Item -LiteralPath $tempRoot -Recurse -Force }
New-Item -ItemType Directory -Path $tempRoot, (Join-Path $tempRoot '_rels'), (Join-Path $tempRoot 'word'), (Join-Path $tempRoot 'word\_rels'), (Join-Path $tempRoot 'docProps') | Out-Null

function Escape-Xml([string]$text) { [System.Security.SecurityElement]::Escape($text) }
function Para([string]$text, [string]$style = 'Normal') {
  $safe = Escape-Xml $text
  return "<w:p><w:pPr><w:pStyle w:val=`"$style`"/></w:pPr><w:r><w:t xml:space=`"preserve`">$safe</w:t></w:r></w:p>"
}
function CodePara([string]$text) {
  $safe = Escape-Xml $text
  return "<w:p><w:pPr><w:pStyle w:val=`"Code`"/></w:pPr><w:r><w:t xml:space=`"preserve`">$safe</w:t></w:r></w:p>"
}
function Bullet([string]$text) {
  $safe = Escape-Xml $text
  return "<w:p><w:pPr><w:pStyle w:val=`"Bullet`"/></w:pPr><w:r><w:t xml:space=`"preserve`">$safe</w:t></w:r></w:p>"
}

$body = New-Object System.Collections.Generic.List[string]
$body.Add((Para 'Comprendre ton code HTML et CSS' 'Title'))
$body.Add((Para 'Guide pédagogique basé sur index.html et asset/styles.css' 'Subtitle'))
$body.Add((Para 'Objectif : expliquer le rôle de chaque partie du code, montrer comment le HTML et le CSS travaillent ensemble, et signaler les points à corriger. Aucun fichier source n’a été modifié.' 'Callout'))

$body.Add((Para '1. Vue d’ensemble du projet' 'Heading1'))
$body.Add((Para 'Ton projet utilise deux fichiers principaux. index.html décrit le contenu et la structure de la page. asset/styles.css contrôle son apparence : couleurs, espacements, alignements et effets visuels.'))
$body.Add((Bullet 'HTML = structure et contenu de la page.'))
$body.Add((Bullet 'CSS = présentation visuelle de cette structure.'))
$body.Add((Bullet 'Font Awesome = bibliothèque externe utilisée pour afficher les icônes des réseaux sociaux.'))
$body.Add((Para 'Le lien entre les deux fichiers est créé dans la partie <head> du document HTML grâce à la balise <link rel="stylesheet" href="asset/styles.css" />.' 'Callout'))

$body.Add((Para '2. Explication de index.html' 'Heading1'))
$htmlItems = @(
  @('<!doctype html>', 'Indique au navigateur que le document utilise HTML5.'),
  @('<html lang="en">', 'Ouvre le document HTML. lang="en" annonce que la langue principale est l’anglais. Pour une page française, la valeur habituelle est fr.'),
  @('<head>', 'Contient les informations destinées au navigateur : encodage, adaptation mobile, titre et feuilles de style.'),
  @('<meta charset="UTF-8" />', 'Active l’encodage UTF-8 afin d’afficher correctement les accents et de nombreux caractères.'),
  @('<meta name="viewport" content="width=device-width, initial-scale=1.0" />', 'Adapte la largeur de la page à celle de l’écran. Cette ligne est essentielle sur téléphone.'),
  @('<title>Brandon Johnson</title>', 'Définit le texte visible dans l’onglet du navigateur.'),
  @('<link rel="stylesheet" href="asset/styles.css" />', 'Charge ta feuille CSS locale. Le chemin signifie : entrer dans le dossier asset puis ouvrir styles.css.'),
  @('Lien vers cdnjs / Font Awesome', 'Charge une feuille CSS externe contenant les icônes utilisées plus bas dans la navigation.'),
  @('<body>', 'Contient tout ce qui doit être visible dans la page.'),
  @('<nav class="navbar">', 'Crée une zone de navigation. La classe navbar permet de la sélectionner en CSS avec .navbar.'),
  @('<div class="nav-links">', 'Regroupe les liens du menu principal afin de les aligner et de les espacer ensemble.'),
  @('<a href="index.html" class="active">Home</a>', 'Crée le lien Home. href indique la destination. La classe active sert à montrer la page actuellement sélectionnée.'),
  @('<a href="about.html">About</a>', 'Crée un lien vers about.html. Le même principe est utilisé pour Resume et Services.'),
  @('<a href="Portfolio.html"></a>', 'Crée un lien vers Portfolio.html, mais le lien est vide : aucun texte ne sera visible à l’écran.'),
  @('<div class="dropdown">', 'Prépare un conteneur pour un menu déroulant. Pour le moment, aucun sous-menu ni comportement déroulant n’est défini.'),
  @('<a href="">Dropdown</a>', 'Affiche Dropdown, mais href est vide : le navigateur reste généralement sur la même page.'),
  @('<a href="">Contact</a>', 'Affiche Contact, mais sans destination réelle pour le moment.'),
  @('<div class="nav-socials">', 'Regroupe les liens des réseaux sociaux, séparément des liens principaux.'),
  @('<a href="#">...</a>', 'Le symbole # sert ici de lien provisoire. Il peut ramener en haut de la page.'),
  @('<i class="fab fa-x-twitter"></i>', 'Demande à Font Awesome d’afficher l’icône X/Twitter. Les autres classes affichent Facebook, Instagram et LinkedIn.'),
  @('</nav>, </body>, </html>', 'Ferment respectivement la navigation, le contenu visible et le document HTML.')
)
foreach ($item in $htmlItems) {
  $body.Add((CodePara $item[0])); $body.Add((Para $item[1]))
}

$body.Add((Para '3. Explication de asset/styles.css' 'Heading1'))
$cssItems = @(
  @('@import url("https://cloudflare.com");', 'Essaie d’importer une autre feuille CSS. Problème : cette adresse pointe vers une page web, pas vers un fichier CSS. Le navigateur ne peut donc pas l’utiliser comme feuille de style.'),
  @('body { ... }', 'Sélectionne tout le contenu visible de la page.'),
  @('margin: 0;', 'Supprime la marge automatique que le navigateur place autour de la page.'),
  @('font-family: sans-serif;', 'Utilise une police sans empattements. Le navigateur choisit une police disponible sur l’appareil.'),
  @('background-color: #0b111e;', 'Applique un bleu très foncé comme arrière-plan général.'),
  @('.navbar { ... }', 'Sélectionne l’élément HTML qui possède class="navbar".'),
  @('display: flex;', 'Active Flexbox. Les enfants directs nav-links et nav-socials peuvent alors être alignés sur une même ligne.'),
  @('justify-content: space-between;', 'Place le premier groupe vers la gauche et le second vers la droite, avec l’espace libre entre eux.'),
  @('align-items: center;', 'Centre verticalement les deux groupes à l’intérieur de la barre.'),
  @('background-color: #040b14;', 'Donne à la barre une couleur légèrement différente du fond de la page.'),
  @('padding: 15px 30px;', 'Ajoute 15 px d’espace intérieur en haut et en bas, puis 30 px à gauche et à droite.'),
  @('border-bottom: 1px solid #101a2a;', 'Ajoute une fine bordure sous la barre : épaisseur 1 px, trait continu, couleur bleu foncé.'),
  @('.nav-links { ... }', 'Sélectionne le conteneur des liens Home, About, Resume, Services, Portfolio, Dropdown et Contact.'),
  @('display: flex;', 'Place les enfants directs du menu sur une ligne.'),
  @('align-items: center;', 'Centre verticalement les liens du menu.'),
  @('gap: 25px;', 'Crée un espace régulier de 25 px entre les éléments du menu.'),
  @('margin-left: 100px;', 'Ajoute une marge extérieure de 100 px à gauche. C’est cette propriété qui pousse ton menu vers la droite.'),
  @('.nav-socials { ... }', 'Sélectionne le groupe des icônes sociales.'),
  @('display: flex; align-items: center;', 'Place les icônes sur une ligne et les centre verticalement.'),
  @('gap: 25px;', 'Ajoute 25 px entre chaque icône.'),
  @('margin-right: 100px;', 'Éloigne le groupe d’icônes du bord droit de 100 px.'),
  @('.nav-links a { ... }', 'Sélectionne tous les liens <a> placés dans nav-links.'),
  @('color: #a8b6cd;', 'Définit une couleur gris bleu clair pour le texte des liens.'),
  @('text-decoration: none;', 'Supprime le soulignement automatique des liens.'),
  @('font-size: 15px;', 'Fixe la taille du texte à 15 pixels.'),
  @('font-weight: 500;', 'Donne au texte une épaisseur moyenne, si la police utilisée possède ce niveau de graisse.'),
  @('transition: color 0.3s ease;', 'Anime les changements de couleur pendant 0,3 seconde avec une accélération et un ralentissement progressifs.'),
  @('position: relative;', 'Crée un repère de positionnement pour le pseudo-élément ::after du lien actif.'),
  @('padding-bottom: 8px;', 'Ajoute 8 px sous le texte afin de laisser de la place à la ligne bleue.'),
  @('.nav-links a:hover { color: #ffff; }', 'Quand la souris passe sur un lien, sa couleur devient blanche. #ffff est une notation RGBA courte valide ; #fff est la notation RGB courte la plus habituelle.'),
  @('.nav-links a.active { color: #149ddd; }', 'Colore en bleu le lien qui possède la classe active.'),
  @('.nav-links a.active::after { ... }', 'Crée un élément graphique virtuel après le contenu du lien actif. Il sert ici de soulignement bleu.'),
  @('content: "";', 'Obligatoire pour afficher ::after. Le pseudo-élément ne contient aucun texte.'),
  @('position: absolute;', 'Permet de placer précisément la ligne par rapport au lien, qui est positionné en relative.'),
  @('bottom: 0; left: 0;', 'Place la ligne en bas du lien et l’aligne sur son bord gauche.'),
  @('width: 100%; height: 2px;', 'Donne à la ligne toute la largeur du lien et une épaisseur de 2 px.'),
  @('background-color: #149ddd;', 'Donne à la ligne la même couleur bleue que le lien actif.')
)
foreach ($item in $cssItems) {
  $body.Add((CodePara $item[0])); $body.Add((Para $item[1]))
}

$body.Add((Para '4. Comment HTML et CSS se correspondent' 'Heading1'))
$body.Add((Para 'En CSS, un point placé avant un nom signifie « sélectionner une classe ». Ainsi, .navbar vise class="navbar", .nav-links vise class="nav-links" et .nav-socials vise class="nav-socials".'))
$body.Add((Para '.nav-links a signifie : sélectionner toutes les balises <a> qui se trouvent à l’intérieur d’un élément portant la classe nav-links. Le sélecteur .nav-links a.active est encore plus précis : il vise seulement un lien qui possède aussi la classe active.' 'Callout'))
$body.Add((Para 'Le mot :hover représente un état temporaire lorsque la souris survole un élément. Le symbole ::after crée un pseudo-élément qui n’existe pas dans le HTML, mais qui peut être dessiné et positionné avec le CSS.'))

$body.Add((Para '5. Points importants à corriger ou compléter' 'Heading1'))
$body.Add((Bullet 'L’import Cloudflare n’est pas une feuille CSS valide et peut produire une erreur dans la console.'))
$body.Add((Bullet 'Le lien Portfolio ne contient aucun texte, donc il est invisible.'))
$body.Add((Bullet 'Le conteneur dropdown ne possède pas encore de sous-menu ni de règles CSS pour l’ouverture et la fermeture.'))
$body.Add((Bullet 'Les destinations vides et les liens # sont seulement provisoires.'))
$body.Add((Bullet 'Les icônes de nav-socials ne possèdent pas encore leurs propres règles de couleur, de décoration ou de survol. Elles peuvent donc garder le style de lien par défaut du navigateur.'))
$body.Add((Bullet 'Les marges fixes de 100 px à gauche et à droite peuvent prendre trop de place sur un petit écran. Une adaptation responsive sera nécessaire pour les téléphones.'))

$body.Add((Para '6. Résumé à retenir' 'Heading1'))
$body.Add((Para 'Ta navigation repose sur deux niveaux de Flexbox : navbar sépare le menu et les réseaux sociaux, tandis que nav-links et nav-socials alignent leurs propres éléments. margin-left pousse le menu principal vers la droite ; margin-right éloigne les réseaux sociaux du bord droit. La classe active et le pseudo-élément ::after créent l’indication bleue de la page courante.' 'Callout'))
$body.Add((Para 'Ce document décrit l’état actuel des fichiers observés. Les explications devront être actualisées si le code change.' 'Small'))

$documentXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
$($body -join "`n")
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1080" w:right="1152" w:bottom="1080" w:left="1152" w:header="708" w:footer="708"/><w:footerReference w:type="default" r:id="rId2" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></w:sectPr>
</w:body></w:document>
"@

$stylesXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="21"/><w:color w:val="26364A"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="21"/><w:color w:val="26364A"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="80" w:after="100"/><w:keepNext/></w:pPr><w:rPr><w:rFonts w:ascii="Aptos Display" w:hAnsi="Aptos Display"/><w:b/><w:sz w:val="52"/><w:color w:val="123B5D"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="300"/></w:pPr><w:rPr><w:sz w:val="25"/><w:color w:val="4E738F"/><w:i/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:spacing w:before="300" w:after="130"/><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:rFonts w:ascii="Aptos Display" w:hAnsi="Aptos Display"/><w:b/><w:sz w:val="31"/><w:color w:val="0F6B78"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="90" w:after="45"/><w:shd w:fill="EAF2F5"/><w:ind w:left="180" w:right="180"/><w:pBdr><w:left w:val="single" w:sz="18" w:space="6" w:color="149DDD"/></w:pBdr></w:pPr><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="18"/><w:color w:val="123B5D"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Callout"><w:name w:val="Callout"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="100" w:after="180"/><w:shd w:fill="F2F8F9"/><w:ind w:left="240" w:right="240"/><w:pBdr><w:left w:val="single" w:sz="20" w:space="8" w:color="0F6B78"/></w:pBdr></w:pPr><w:rPr><w:color w:val="244A55"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Bullet"><w:name w:val="Bullet"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="70"/><w:ind w:left="420" w:hanging="220"/></w:pPr><w:rPr><w:color w:val="26364A"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Small"><w:name w:val="Small"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="18"/><w:color w:val="687785"/><w:i/></w:rPr></w:style>
</w:styles>
'@

$footerXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:color w:val="7A8894"/><w:sz w:val="17"/></w:rPr><w:t>Guide HTML &amp; CSS — Brandon Johnson</w:t></w:r></w:p></w:ftr>
'@

$contentTypes = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>
'@
$rootRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>
'@
$docRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>
'@
$coreXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Comprendre ton code HTML et CSS</dc:title><dc:subject>Guide pédagogique</dc:subject><dc:creator>Codex</dc:creator><cp:keywords>HTML, CSS, navigation, Flexbox</cp:keywords><dcterms:created xsi:type="dcterms:W3CDTF">2026-08-29T00:00:00Z</dcterms:created></cp:coreProperties>
'@
$appXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Office Word</Application></Properties>
'@

$utf8 = New-Object System.Text.UTF8Encoding($false)
[IO.File]::WriteAllText((Join-Path $tempRoot '[Content_Types].xml'), $contentTypes, $utf8)
[IO.File]::WriteAllText((Join-Path $tempRoot '_rels\.rels'), $rootRels, $utf8)
[IO.File]::WriteAllText((Join-Path $tempRoot 'word\document.xml'), $documentXml, $utf8)
[IO.File]::WriteAllText((Join-Path $tempRoot 'word\styles.xml'), $stylesXml, $utf8)
[IO.File]::WriteAllText((Join-Path $tempRoot 'word\footer1.xml'), $footerXml, $utf8)
[IO.File]::WriteAllText((Join-Path $tempRoot 'word\_rels\document.xml.rels'), $docRels, $utf8)
[IO.File]::WriteAllText((Join-Path $tempRoot 'docProps\core.xml'), $coreXml, $utf8)
[IO.File]::WriteAllText((Join-Path $tempRoot 'docProps\app.xml'), $appXml, $utf8)

Add-Type -AssemblyName System.IO.Compression.FileSystem
if (Test-Path -LiteralPath $outputPath) { Remove-Item -LiteralPath $outputPath -Force }
[System.IO.Compression.ZipFile]::CreateFromDirectory($tempRoot, $outputPath)
Remove-Item -LiteralPath $tempRoot -Recurse -Force
Write-Output $outputPath
