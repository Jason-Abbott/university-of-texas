<%
' this bit of script allows you to set a color
' theme for webCal
' un-comment the theme you want to activate
' and comment the old theme

dim arColor

' the first seven are generally different shades of the
' same color from light to dark, each value decremented
' by Hex 10

' Grayscale:
'arColor = Array("e0e0e0","d0d0d0","c0c0c0","b0b0b0","a0a0a0","909090","808080","303030","e0e0e0","ffffff","a0a0a0","e0e0e0","c0c0c0","a0a0a0","808080")

' Turqoise on white background:
'arColor = Array("dffcfc","ffffff","cfecec","bfdcdc","afcccc","9fbcbc","8facac","003355","ff7777","f0f0ff","99bbdd","e0e0e0","c0c0c0","a0a0a0","808080")

' Spring colors:
'arColor = Array("a9ff76","99ff66","89df56","79cf46","69bf36","59af26","499f16","004400","ffff44","ffffff","ffcc33","e0e0e0","c0c0c0","a0a0a0","808080")

' Brown colors:
'arColor = Array("ffdca9","ffcc99","efbc89","dfac79","cf9c69","bf8c59","af7c49","661100","00ccff","ffffff","cc9966","e0e0e0","c0c0c0","a0a0a0","808080")

' Default purplish colors:
arColor = Array("efedff","dfddff","cfcdef","bfbddf","afadcf","9f9dbf","8f8daf","330033","ffcc33","ffffff","c6c6d9","e0e0e0","c0c0c0","a0a0a0","808080")
'arColor = Array("75a8ff","5588ff","3568df","1548bf","ffffff","02087f","00045f","330033","ffcc33","ffffff","c6c6d9","e0e0e0","c0c0c0","a0a0a0","808080")
'              0        1        2        3        4        5        6        7        8        9        10       11       12       13       14

' here's a key to the colors as presently used:
' 0 =             [lightest shade]
' 1 = background  .
' 2 =             . 
' 3 =             .
' 4 = title text, heading background
' 5 =             .
' 6 = active link [darkest shade]
' 7 = link        [begin irregular colors]
' 8 = today
' 9 = weekday
'10 = weekend
'[the following probably won't need to be changed]
'11 = non-day back, form light  [light grey]
'12 =                           .
'13 = non-day text              .
'14 = form text                 [dark grey]
%>

<style type="text/css">
	.child { font-size:10; display:block; }
	A:hover { color:#<%=arColor(6)%>; }
	A:link { text-decoration: none; }
	A:active { text-decoration: none; }
	A:visited { text-decoration: none; }
</style>