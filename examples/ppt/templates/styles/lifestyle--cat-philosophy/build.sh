#!/bin/bash
set -e

FILE="cat_philosophy.pptx"
echo "Creating PPTX: $FILE"
officecli create "$FILE"

echo "Adding Slide 1 (hero)..."
officecli add "$FILE" '/' --type slide --prop layout=blank --prop background=FFF9E6
echo '[
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"circle-main","preset":"ellipse","fill":"FF8A4C","x":"18cm","y":"3cm","width":"18cm","height":"18cm","opacity":"1.0"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"circle-sub","preset":"ellipse","fill":"FFC533","x":"26cm","y":"12cm","width":"10cm","height":"10cm","opacity":"1.0"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"round-box","preset":"roundRect","fill":"FFC533","x":"5cm","y":"12cm","width":"12cm","height":"6cm","rotation":"-10","opacity":"0.3"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"line-top","preset":"roundRect","fill":"4A3B32","x":"3cm","y":"2cm","width":"6cm","height":"0.4cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"dot-small","preset":"ellipse","fill":"4A3B32","x":"28cm","y":"3cm","width":"1.5cm","height":"1.5cm"}},

  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"s1-title","text":"猫咪的统治哲学","font":"Source Han Sans","size":"64","bold":"true","color":"2A201A","x":"3cm","y":"4cm","width":"22cm","height":"4cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"s1-sub","text":"为什么地球人自愿成为“铲屎官”？","font":"Source Han Sans","size":"36","color":"4A3B32","x":"3cm","y":"8.5cm","width":"24cm","height":"2cm"}},
  {"command":"add","parent":"/slide[1]","type":"shape","props":{"name":"s1-tag","text":"一场长达一万年的双向奔赴","font":"Source Han Sans","size":"20","color":"FF8A4C","bold":"true","x":"3cm","y":"11.5cm","width":"15cm","height":"1.5cm"}}
]' | officecli batch "$FILE"


echo "Adding Slide 2 (statement)..."
officecli add "$FILE" '/' --from '/slide[1]'
echo '[
  {"command":"set","path":"/slide[2]","props":{"transition":"morph"}},
  {"command":"set","path":"/slide[2]/shape[1]","props":{"x":"5cm","y":"0cm","width":"26cm","height":"26cm","opacity":"0.1"}},
  {"command":"set","path":"/slide[2]/shape[2]","props":{"x":"10cm","y":"10cm","width":"18cm","height":"18cm","opacity":"0.1"}},
  {"command":"set","path":"/slide[2]/shape[3]","props":{"x":"36cm","y":"5cm"}},
  {"command":"set","path":"/slide[2]/shape[4]","props":{"x":"14cm","y":"4cm","width":"8cm"}},
  {"command":"set","path":"/slide[2]/shape[5]","props":{"x":"6cm","y":"14cm","width":"2cm","height":"2cm"}},
  {"command":"set","path":"/slide[2]/shape[6]","props":{"x":"36cm"}},
  {"command":"set","path":"/slide[2]/shape[7]","props":{"x":"36cm"}},
  {"command":"set","path":"/slide[2]/shape[8]","props":{"x":"36cm"}},

  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"s2-title","text":"这不是你养了宠物，\\n这是一场完美的跨物种PUA。","font":"Source Han Sans","size":"54","bold":"true","color":"2A201A","align":"center","x":"4cm","y":"6cm","width":"26cm","height":"5cm"}},
  {"command":"add","parent":"/slide[2]","type":"shape","props":{"name":"s2-sub","text":"狗被驯化用来工作，而猫走进人类生活，只因为这里有免费的食物和暖炉。","font":"Source Han Sans","size":"24","color":"4A3B32","align":"right","x":"12cm","y":"13cm","width":"18cm","height":"3cm"}}
]' | officecli batch "$FILE"


echo "Adding Slide 3 (pillars)..."
officecli add "$FILE" '/' --from '/slide[2]'
echo '[
  {"command":"set","path":"/slide[3]","props":{"transition":"morph"}},
  {"command":"set","path":"/slide[3]/shape[1]","props":{"x":"0cm","y":"12cm","width":"12cm","height":"12cm","opacity":"0.2"}},
  {"command":"set","path":"/slide[3]/shape[2]","props":{"x":"36cm","y":"0cm"}},
  {"command":"set","path":"/slide[3]/shape[3]","props":{"x":"2cm","y":"2cm","width":"8cm","height":"8cm","opacity":"0.1","rotation":"0"}},
  {"command":"set","path":"/slide[3]/shape[4]","props":{"x":"2cm","y":"2cm","width":"8cm"}},
  {"command":"set","path":"/slide[3]/shape[5]","props":{"x":"30cm","y":"2cm","width":"3cm","height":"3cm"}},
  {"command":"set","path":"/slide[3]/shape[9]","props":{"x":"36cm"}},
  {"command":"set","path":"/slide[3]/shape[10]","props":{"x":"36cm"}},

  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"s3-title","text":"统治地球的三大核心武器","font":"Source Han Sans","size":"44","bold":"true","color":"2A201A","x":"2cm","y":"3cm","width":"24cm","height":"2.5cm"}},

  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"p1-bg","preset":"roundRect","fill":"FFFFFF","x":"2cm","y":"6.5cm","width":"9cm","height":"10.5cm","opacity":"0.8","animation":"fade-entrance-400-with"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"p2-bg","preset":"roundRect","fill":"FFFFFF","x":"12.5cm","y":"6.5cm","width":"9cm","height":"10.5cm","opacity":"0.8","animation":"fade-entrance-400-with-delay=100"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"p3-bg","preset":"roundRect","fill":"FFFFFF","x":"23cm","y":"6.5cm","width":"9cm","height":"10.5cm","opacity":"0.8","animation":"fade-entrance-400-with-delay=200"}},

  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"p1-title","text":"① 幼态延续","font":"Source Han Sans","size":"26","bold":"true","color":"FF8A4C","x":"3cm","y":"7.5cm","width":"7cm","height":"1.5cm","animation":"fade-entrance-400-with"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"p1-desc","text":"大眼睛、小鼻子，触发人类的本能抚育冲动。Baby Schema 让人类无法抗拒。","font":"Source Han Sans","size":"18","color":"4A3B32","x":"3cm","y":"9.5cm","width":"7cm","height":"5cm","animation":"fade-entrance-400-with"}},

  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"p2-title","text":"② 专属夹子音","font":"Source Han Sans","size":"26","bold":"true","color":"FF8A4C","x":"13.5cm","y":"7.5cm","width":"7cm","height":"1.5cm","animation":"fade-entrance-400-with-delay=100"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"p2-desc","text":"成年猫之间不喵喵叫。这种特定频率专门用来模拟婴儿啼哭，精准操控人类神经。","font":"Source Han Sans","size":"18","color":"4A3B32","x":"13.5cm","y":"9.5cm","width":"7cm","height":"5cm","animation":"fade-entrance-400-with-delay=100"}},

  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"p3-title","text":"③ 间歇性强化","font":"Source Han Sans","size":"26","bold":"true","color":"FF8A4C","x":"24cm","y":"7.5cm","width":"7cm","height":"1.5cm","animation":"fade-entrance-400-with-delay=200"}},
  {"command":"add","parent":"/slide[3]","type":"shape","props":{"name":"p3-desc","text":"时而高冷，时而粘人。在心理学上，这是最容易让人上瘾的反馈机制。","font":"Source Han Sans","size":"18","color":"4A3B32","x":"24cm","y":"9.5cm","width":"7cm","height":"5cm","animation":"fade-entrance-400-with-delay=200"}}
]' | officecli batch "$FILE"


echo "Adding Slide 4 (evidence)..."
officecli add "$FILE" '/' --from '/slide[3]'
echo '[
  {"command":"set","path":"/slide[4]","props":{"transition":"morph"}},
  {"command":"set","path":"/slide[4]/shape[1]","props":{"x":"15cm","y":"0cm","width":"26cm","height":"26cm","opacity":"1.0"}},
  {"command":"set","path":"/slide[4]/shape[2]","props":{"x":"36cm","y":"0cm"}},
  {"command":"set","path":"/slide[4]/shape[3]","props":{"x":"24cm","y":"8cm","width":"12cm","height":"12cm","opacity":"1.0","rotation":"15"}},
  {"command":"set","path":"/slide[4]/shape[4]","props":{"x":"2cm","y":"3cm","width":"5cm"}},
  {"command":"set","path":"/slide[4]/shape[5]","props":{"x":"36cm","y":"0cm"}},
  {"command":"set","path":"/slide[4]/shape[11]","props":{"x":"36cm"}},
  {"command":"set","path":"/slide[4]/shape[12]","props":{"x":"36cm"}},
  {"command":"set","path":"/slide[4]/shape[13]","props":{"x":"36cm"}},
  {"command":"set","path":"/slide[4]/shape[14]","props":{"x":"36cm"}},
  {"command":"set","path":"/slide[4]/shape[15]","props":{"x":"36cm"}},
  {"command":"set","path":"/slide[4]/shape[16]","props":{"x":"36cm"}},
  {"command":"set","path":"/slide[4]/shape[17]","props":{"x":"36cm"}},
  {"command":"set","path":"/slide[4]/shape[18]","props":{"x":"36cm"}},
  {"command":"set","path":"/slide[4]/shape[19]","props":{"x":"36cm"}},
  {"command":"set","path":"/slide[4]/shape[20]","props":{"x":"36cm"}},

  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"s4-title","text":"不仅控制心，还控制多巴胺","font":"Source Han Sans","size":"40","bold":"true","color":"2A201A","x":"2cm","y":"4cm","width":"15cm","height":"2.5cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"s4-data1-val","text":"25-150","font":"Montserrat","size":"80","bold":"true","color":"FFFFFF","align":"right","x":"15cm","y":"5cm","width":"13cm","height":"4cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"s4-data1-unit","text":"Hz","font":"Montserrat","size":"40","bold":"true","color":"FFFFFF","x":"28cm","y":"7.5cm","width":"4cm","height":"2cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"s4-data1-desc","text":"猫咪呼噜声频率，医学证明能促进骨骼愈合","font":"Source Han Sans","size":"20","color":"FFFFFF","align":"right","x":"18cm","y":"10.5cm","width":"14cm","height":"2cm"}},
  
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"s4-data2-title","text":"降低皮质醇","font":"Source Han Sans","size":"28","bold":"true","color":"FF8A4C","x":"2cm","y":"8cm","width":"12cm","height":"1.5cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"s4-data2-desc","text":"看猫咪视频能瞬间降低压力荷尔蒙，提升多巴胺。","font":"Source Han Sans","size":"18","color":"4A3B32","x":"2cm","y":"9.5cm","width":"12cm","height":"3cm"}},

  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"s4-data3-title","text":"弓形虫假说","font":"Source Han Sans","size":"28","bold":"true","color":"FF8A4C","x":"2cm","y":"13cm","width":"12cm","height":"1.5cm"}},
  {"command":"add","parent":"/slide[4]","type":"shape","props":{"name":"s4-data3-desc","text":"猫咪体内的寄生虫可能悄悄改变了人类的冒险神经。","font":"Source Han Sans","size":"18","color":"4A3B32","x":"2cm","y":"14.5cm","width":"12cm","height":"3cm"}}
]' | officecli batch "$FILE"


echo "Adding Slide 5 (cta)..."
officecli add "$FILE" '/' --from '/slide[4]'
echo '[
  {"command":"set","path":"/slide[5]","props":{"transition":"morph"}},
  {"command":"set","path":"/slide[5]/shape[1]","props":{"x":"12cm","y":"4cm","width":"10cm","height":"10cm","opacity":"0.8"}},
  {"command":"set","path":"/slide[5]/shape[2]","props":{"x":"8cm","y":"3cm","width":"6cm","height":"6cm","opacity":"1.0"}},
  {"command":"set","path":"/slide[5]/shape[3]","props":{"x":"36cm","y":"0cm"}},
  {"command":"set","path":"/slide[5]/shape[4]","props":{"x":"14cm","y":"15cm","width":"6cm"}},
  {"command":"set","path":"/slide[5]/shape[5]","props":{"x":"16cm","y":"5cm","width":"2cm","height":"2cm"}},
  {"command":"set","path":"/slide[5]/shape[21]","props":{"x":"36cm"}},
  {"command":"set","path":"/slide[5]/shape[22]","props":{"x":"36cm"}},
  {"command":"set","path":"/slide[5]/shape[23]","props":{"x":"36cm"}},
  {"command":"set","path":"/slide[5]/shape[24]","props":{"x":"36cm"}},
  {"command":"set","path":"/slide[5]/shape[25]","props":{"x":"36cm"}},
  {"command":"set","path":"/slide[5]/shape[26]","props":{"x":"36cm"}},
  {"command":"set","path":"/slide[5]/shape[27]","props":{"x":"36cm"}},
  {"command":"set","path":"/slide[5]/shape[28]","props":{"x":"36cm"}},

  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"s5-title","text":"接受现实吧","font":"Source Han Sans","size":"64","bold":"true","color":"2A201A","align":"center","x":"4cm","y":"6cm","width":"26cm","height":"4cm"}},
  {"command":"add","parent":"/slide[5]","type":"shape","props":{"name":"s5-sub","text":"今天你给主子开罐头了吗？","font":"Source Han Sans","size":"32","color":"4A3B32","align":"center","x":"4cm","y":"11cm","width":"26cm","height":"2cm"}}
]' | officecli batch "$FILE"

echo "Done."