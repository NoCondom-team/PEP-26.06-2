#!/usr/bin/env python3
"""Build Romanian HTML content from text.docx structure."""

import re
import zipfile
import xml.etree.ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

IMAGE_FILES = {
    1: 'pic1.jpg', 2: 'pic2.jpg', 3: 'pic3.png', 4: 'pic4.jpg',
    5: 'pic5.png', 6: 'pic6.jpg', 7: 'pic7.webp', 8: 'pic8.jpg',
    9: 'pic9.avif', 10: 'pic10.jpg', 11: 'pic11.jpg', 12: 'pic12.jpg',
}

HIGHLIGHT_COLORS = {
    'green': '#00FF00', 'yellow': '#FFFF00', 'red': '#FF0000',
    'cyan': '#00FFFF', 'magenta': '#FF00FF',
}

# Romanian translations keyed by paragraph index (from docx)
RO = {
    0: '«Mi-au mai rămas 3 luni. M-a ucis cistita pe care am tratat-o greșit timp de 10 ani. Citiți asta înainte să fie prea târziu»',
    1: 'Scrisoarea Elenei Matei, 62 de ani, Brașov. Publicată la cererea ei.',
    2: 'Redacția a primit această scrisoare pe 12 martie. Elena a încetat din viață pe 18 martie. A cerut să o publicăm, pentru ca alte femei să nu repete drumul ei. Îndeplinim ultima ei dorință.',
    4: '„Mă numesc Elena. Am 62 de ani. Când citiți asta, cel mai probabil nu mai sunt.',
    5: 'Medicii mi-au dat trei luni. Insuficiență renală. Stadiu terminal. Dializa nu ajută — organismul este prea epuizat. Stau în apartamentul meu din Brașov și scriu această scrisoare de mână, pentru că nu vreau să duc adevărul cu mine în mormânt.',
    6: 'Nu m-a ucis cistita. M-a ucis ignoranța. A mea și a medicilor.”',
    8: 'Cum a început totul',
    9: 'Aveam 52 de ani. Menopauză. Flăcări noaptea — mă trezeam udă, ca după duș. Treptat am luat în greutate — în doi ani am pus 8 kg, deși mâncam ca de obicei. A apărut slăbiciune după masă — atât de mare, că adormeam la volan. Uscăciune în gură dimineața — mă trezeam cu senzația că am deșert în gură. Și încă — îmi dorea mereu dulce. Nu puteam trece pe lângă cofetărie. Ciocolată, biscuiți, chiflă — asta devenise mângâierea mea.',
    10: 'Atunci a început și cistita. Mai întâi o dată la șase luni. Apoi mai des. Imagine tipică: dimineața mă trezesc, merg la toaletă, și brusc — arsură. Înțepătură. Senzația că prin uretră trece sârmă încinsă. Urina tulbure, cu miros puternic — nu ca de obicei, ci dulceag, grețos. Mi-am amintit asta pentru totdeauna. Uneori — roz de la sânge.',
    12: 'Urologul de la policlinică mi-a prescris Monural. Un plic seara. A ajutat o lună. Am răsuflat ușurată. Dar după o lună — din nou. Aceeași arsură, aceleași înțepături, același miros dulceag al urinei. Alt urolog a prescris Ciprol. A ajutat două săptămâni. Al treilea — Fosfomicină. A ajutat cinci zile. Al patrulea și-a ridicat mâinile și a spus: «Ei bine, ce vreți, vârsta, menopauza. Beți suc de merișoare și învățați să trăiți cu asta».',
    13: 'Am mers în cerc timp de 10 ani. Antibiotice, picături din plante, suc de merișoare cu litrii, tampoane urologice. Cheltuiam pe toate acestea peste 3.000 de lei pe an. În 10 ani — 30.000 de lei. Am dat acești bani ca să mor încet.',
    15: 'Simptomele pe care le-am ignorat',
    16: 'Acum, privind înapoi, văd cum totul se lega. Dar atunci nu conectam una cu alta.',
    17: 'Oboseală constantă. Nu doar «obosită după serviciu». Ci atât de mare, că nu puteam să mă ridic din pat. Soțul spunea: «Ești pur și simplu leneșă». Iar eu nu aveam putere să explic că nu e vorba de lene.',
    18: 'Poftă de dulce. Puteam mânca o întreagă tabletă de ciocolată într-o seară. Și tot nu mă săturam. Organismul striga: «Dă-mi zahăr!». Iar eu dădeam. Și mă omoram.',
    19: 'Uscăciune în gură. Mai ales dimineața. Limba ca șmirghel. Buzele crăpau până la sânge.',
    20: 'Mâncărime a pielii. Mâncau mâinile, picioarele, spatele. Mai ales noaptea. Schimbam detergent, mă ungeam cu creme — nimic nu ajuta.',
    21: 'Agravarea vederii. Îmi schimbam ochelarii la fiecare șase luni. Literele se estompau. Credeam — vârsta. De fapt, zahărul distrugea vasele de sânge ale ochilor.',
    22: 'Răni care se vindecă greu. M-am tăiat la deget cu cuțitul — s-a vindecat în trei săptămâni. Credeam — ei bine, vârsta, metabolismul s-a încetinit.',
    23: 'Urgențe frecvente la toaletă noaptea. Nu o dată, nu de două ori. Trei, patru, cinci ori pe noapte. Nu dormeam de ani. Mă trezeam la 2:15, la 3:40, la 5:10. Urina tulbure, cu miros dulceag puternic. Dimineața — distrusă, ca și cum n-aș fi dormit deloc.',
    24: 'Și cistita. Constantă. Care se întorcea.',
    25: 'Niciun medic nu a legat totul într-o singură imagine. Niciunul nu a spus: «Hai să verificăm insulina». Niciunul nu a explicat că mirosul urinei și pofta de dulce sunt verigi ale aceleiași lanț.',
    27: 'Ce am aflat prea târziu',
    28: 'La 61 de ani am ajuns la endocrinolog. Din întâmplare. Făceam control medical pentru sanatoriu. Medicul s-a uitat la analizele mele și a întrebat: «Știți că aveți prediabet?»',
    31: 'Prediabet. Rezistență ascunsă la insulină. Insulina mea era de 4 ori mai mare decât normal. Organismul o producea cu tone, ca să țină zahărul în sânge. Dar zahărul tot se scurgea în urină. Urina mea era dulce. Literalmente. Bacteriile din vezica urinară primeau mediu nutritiv ideal. Cald, umed, zahăr — bufet suedez non-stop. Beam antibiotic — ucidea bacteriile. Dar zahărul rămânea. Și după o săptămână creștea o colonie nouă.',
    32: 'E ca și cum ai încerca să scoți apa dintr-o barcă cu gaură. Scoți, iar apa tot vine.',
    33: 'Am întrebat: «De ce nimeni nu mi-a spus asta mai devreme? De ce niciun urolog nu a verificat insulina?»',
    34: 'Endocrinologul a răspuns: «Pentru că urologia și endocrinologia sunt specialități diferite. Urologul caută bacterii în urină. Endocrinologul caută probleme hormonale. Iar ce e între ele — teritoriu nimănui. Ați ajuns în zona moartă a medicinei».',
    35: 'În zona moartă. Din care nu am mai ieșit.',
    36: 'Până atunci era prea târziu. Cistita a trecut în pielonefrită. Pielonefrită — în insuficiență renală cronică. Iar rezistența la insulină în acești 10 ani s-a transformat în diabet de tip 2 adevărat.',
    37: 'Zahărul și infecția mi-au ucis rinichii.',
    41: 'Cum funcționează — explicația pe care ar fi trebuit să o aud acum 10 ani',
    42: 'Acum știu. Dar prea târziu.',
    43: 'Când aveți rezistență la insulină, pancreasul produce prea multă insulină ca să împingă glucoza în celule. Celulele rezistă. Insulina tot crește. Și la un moment dat rinichii încep să lase glucoza în urină.',
    44: 'Urina devine dulce. Analiza obișnuită nu arată asta — glucoza în probe unice e adesea normală. Dar dacă colectezi urina de 24 de ore — acolo e zahăr.',
    45: 'Bacteriile care trăiesc în tractul urinar primesc hrană nesfârșită. Se înmulțesc, formează biofilm — strat mucos care le protejează de antibiotice. Bei antibiotic — stratul superior de bacterii moare, devine mai ușor. Dar biofilmul rămâne. Și după o săptămână eliberează o armată nouă.',
    46: 'Iar zahărul între timp distruge vasele de sânge. Pereții vezicii urinare nu primesc suficient oxigen. Mucoasa se subțiază. Apar microfisuri. Bacteriile pătrund adânc în țesuturi. Începe inflamația cronică.',
    47: 'An după an. Antibiotic după antibiotic. Până când rinichii spun «destul».',
    49: 'Ce a creat doctorul Popescu pentru oameni ca mine',
    50: 'Cu o lună înainte să mi se pună diagnosticul terminal, am ajuns la consultația doctorului Irnel Popescu. Singurul din România care combină endocrinologia și urologia. S-a uitat la analizele mele vechi, a făcut test extins de insulină și a spus:',
    52: '«Elena, cistita dumneavoastră nu a fost niciodată o infecție. A fost un simptom al rezistenței la insulină. Dacă acum 10 ani vi s-ar fi prescris berberină și acid alfa-lipoic, acum ați fi sănătoasă».',
    53: 'Am întrebat: «De ce nimeni nu mi-a spus asta mai devreme?»',
    54: 'A răspuns: «Pentru că nu e profitabil. Să calculăm împreună, Elena. Cheltuiați pe antibiotice, supozitoare, tampoane și analize 3.000 de lei pe an. În 10 ani — 30.000 de lei. Și nu erați singură. În România peste 2 milioane de femei după 45 de ani suferă de cistită cronică recurentă. Dacă fiecare cheltuie cel puțin 2.500 de lei pe an — iar multe cheltuie mai mult — asta înseamnă 5 miliarde de lei pe an. Cinci miliarde doar în țara noastră. Doar pe cistită. Doar pe ceea ce nu funcționează.',
    55: 'Și acum imaginați-vă că fiecăreia i se spune adevărul: «Problema dumneavoastră nu e în vezica urinară. Problema e în insulină. Luați berberină și acid alfa-lipoic. Costă 111 lei pe curs, o dată pe an.» Asta înseamnă 222 de milioane de lei în loc de 5 miliarde. Pierdere de 4,8 miliarde de lei pe an pentru industria farmaceutică. Doar în România.',
    56: 'Iar la scară mondială? De cistită cronică suferă peste 400 de milioane de femei. Piața de antibiotice, tampoane urologice, antiinflamatoare și «ceaiuri din plante» pentru ele — 200 de miliarde de euro pe an. Și toți acești 200 de miliarde se prăbușesc la zero, dacă femeile află că cauza cistitei nu sunt bacteriile, ci zahărul.',
    57: 'De aceea cercetările despre rezistența la insulină și cistită nu sunt finanțate. De aceea la conferințele medicale urologilor nu li se spune despre indicele HOMA-IR. De aceea în farmacii vi se oferă Monural la 45 de lei, nu berberină la câțiva bănuți. Cistita dumneavoastră hrănea o întreagă industrie. O industrie care câștigă pe durerea dumneavoastră. Pe ignoranța dumneavoastră. Pe moartea dumneavoastră, Elena. Și pe moartea miilor ca dumneavoastră».',
    58: 'Doctorul Popescu a spus că, după ce a văzut zeci de cazuri ca al meu, nu mai putea dormi liniștit. S-a închis în laborator și a început să lucreze.',
    59: '«Am răsfoit 1.200 de articole științifice. Căutam nu chimie, nu sintetică, nu încă un antibiotic. Căutam componente naturale care pot face ceea ce nu face niciun urolog: să scoată zahărul din urină și să restabilească sensibilitatea celulelor la insulină. Am petrecut 14 luni în laborator. Am testat formule pe mine. Am schimbat proporțiile. Am testat din nou. Și când am găsit raportul perfect, am înțeles: asta e.',
    60: 'Trei componente. Trei chei pentru o singură problemă. Dar doar în dozaj exact — cu un miligram în plus sau în minus, efectul e ca la placebo.',
    61: 'Primul: berberină. Extract de berberis. 500 mg într-o capsulă. În formă pură, fără impurități. Funcționează ca metformin natural — cu o diferență: metforminul distruge ficatul la administrare prelungită, iar berberina — nu. Reduce rezistența la insulină și scoate zahărul din urină. Bacteriile își pierd baza alimentară în 72 de ore. Pur și simplu nu se mai înmulțesc. Nu au ce mânca.',
    63: 'Al doilea: acid alfa-lipoic. 300 mg. Cel mai puternic antioxidant, care acționează la nivel celular. «Repară» receptorii celulari deteriorați de insulină ridicată. Imaginați-vă o încuietoare acoperită cu clei. Cheia nu intră. La fel și glucoza nu poate intra în celulă când receptorii sunt deteriorați. Acidul alfa-lipoic dizolvă acest «clei». Celulele se deschid din nou pentru glucoză. Zahărul pleacă din sânge în celule — și nu ajunge în urină.',
    65: 'Al treilea: extract de scorțișoară de Ceylon. 200 mg. Nu cassia maroie de la supermarket — în ea e de 250 de ori mai multă cumarină, și distruge ficatul. Scorțișoara de Ceylon — sigură, cu conținut de cumarină sub 0,004%. Imită acțiunea insulinei și ajută glucoza să intre în celule fără creșteri bruște de hormon. Ușor, lin, fără încărcare pe pancreas.',
    67: 'Am amestecat aceste trei componente într-o singură formulă. Am testat pe 500 de femei. La 82% cistita a dispărut complet în 30 de zile. Fără antibiotice. Fără recidive. Am numit-o Cystiolla. Pentru că nu e doar un remediu pentru cistită. E un remediu pentru cauza ei».',
    68: '«Pentru dumneavoastră, Elena, e deja prea târziu, — a spus el. — Dar pentru mii de alte femei — încă nu».',
    70: 'Ultima mea rugăminte',
    71: 'Nu am apucat să încerc Cystiolla. Rinichii mei erau deja morți. Dar am rugat doctorul Popescu să-mi dea contactele femeilor care au apucat. Voiam să aud că măcar cineva s-a salvat.',
    72: 'Iată ce mi-au spus.',
    73: 'Mariana, 53 de ani, Cluj:\n«Elena, plâng citind scrisoarea dumneavoastră. Am avut aceeași poveste — cistită ani de zile, nimeni nu putea ajuta. Nici eu nu puteam trăi fără dulce. Mă trezeam de 5 ori pe noapte. Urina avea miros dulceag puternic. Doctorul Popescu a găsit rezistența la insulină și mi-a dat Cystiolla. Au trecut două luni — niciun atac. Zahărul din sânge s-a normalizat. Nu mă mai trezesc noaptea. Dorm 7 ore la rând. Trăiesc. Trăiesc pentru că am aflat adevărul mai devreme decât dumneavoastră».',
    74: 'Doina, 61 de ani, Iași:\n«Elena, și eu aveam uscăciune în gură dimineața și poftă de dulce. Și cistită cu sânge. Analizele arătau bacterii, antibioticele nu ajutau. Cystiolla a schimbat totul. În două săptămâni urina a devenit limpede, fără miros. Într-o lună analizele erau sterile. Și știți ce e cel mai uimitor? Nu mi-a mai fost dor de dulce. Pur și simplu nu mai voiam. Organismul nu mai cerea zahăr. Sunt sănătoasă. Iar dumneavoastră nu mai sunteți. Cât de nedrept».',
    75: 'Loredana, 49 de ani, Timișoara:\n«Lucrez la birou, gustări constante cu biscuiți. Cistita era în fiecare lună. Pielea mânca, mai ales noaptea. Rănile se vindecau două săptămâni. Atribuiam totul stresului. Doctorul Popescu a spus: aveți rezistență la insulină, și ea provoacă cistita. Cystiolla și renunțarea la dulce au schimbat totul. Cistita a dispărut. Pielea nu mai mânca. Am slăbit 5 kg. Citesc scrisoarea dumneavoastră și înțeleg: puteam fi eu».',
    77: 'Verificați-vă chiar acum',
    78: 'Nu sunt medic. Dar am parcurs acest drum până la capăt. Și vreau să vă opriți și să vă gândiți.',
    79: 'Dacă aveți peste 45 de ani și aveți cel puțin trei dintre aceste semne — cel mai probabil aveți rezistență ascunsă la insulină, și anume ea provoacă cistita:',
    80: 'Oboseală constantă, mai ales după masă — ochii se închid, vreți să dormiți.',
    81: 'Poftă de dulce — ciocolată, biscuiți, chifle. Nu puteți trece pe lângă vitrina cu deserturi.',
    82: 'Uscăciune în gură dimineața — limba ca hârtie abrazivă, buzele se crăpă.',
    83: 'Mâncărime a pielii — mai ales seara și noaptea, fără erupție vizibilă.',
    84: 'Greutate în plus în zona abdomenului — grăsimea se depune la talie, ca o centură de salvare.',
    85: 'Agravarea vederii — literele se estompază, trebuie să schimbați des ochelarii.',
    86: 'Răni și vânătăi care se vindecă greu.',
    87: 'Urgențe frecvente la toaletă noaptea — de două, trei, cinci ori pe noapte.',
    88: 'Urină cu miros dulceag puternic — ați observat asta, dar v-a fost rușine să spuneți cuiva.',
    89: 'Cistită care se întoarce iar și iar, în ciuda antibioticelor.',
    90: 'Dacă v-ați recunoscut — opriți-vă. Mergeți pe drumul meu.',
    93: 'Nu fiți ca mine',
    94: 'Doctorul Popescu a spus că prețul obișnuit al unui curs Cystiolla este 450 de lei. La început am gândit: «Scump». Apoi mi-a pus cifrele pe masă. Și am tăcut.',
    96: '«Uitați-vă, Elena. 450 de lei — asta e prețul unui singur curs. Unu singur. Și acum amintiți-vă cât cheltuiați în fiecare an. Monural — 45 de lei pe plic, de 6 ori pe an — 270 de lei. Ciprol — 35 de lei pe cutie, de 4 ori pe an — 140 de lei. Supozitoare pentru candidoză după fiecare antibiotic — 60 de lei pe cutie, de 6 ori pe an — 360 de lei. Tampoane urologice — 25 de lei pe pachet, 4 pachete pe lună, 12 luni — 1.200 de lei. Canephron, Phytolysin, frunze de merișor — 40 de lei pe cutie, în fiecare lună — 480 de lei. Vizite la urolog, analize de urină, ecografie — minim 500 de lei pe an. Total — aproape 3.000 de lei pe an. Și așa 10 ani la rând. 30.000 de lei pentru 10 ani de durere. Și nici o zi fără gândul la toaletă.',
    97: 'Iar Cystiolla — 450 de lei. Un curs. O dată pe an. Fără antibiotice. Fără tampoane. Fără analize la fiecare două luni. 450 de lei față de 3.000. Asta înseamnă o economie de 2.550 de lei pe an. În 10 ani — 25.500 de lei care rămân în buzunarul dumneavoastră, nu în casa de marcat a farmaciei».',
    98: 'Mă uitam la aceste cifre și nu puteam crede. Am dat 30.000 de lei ca să mor. Iar puteam da 450 — și să trăiesc.',
    99: 'Dar doctorul Popescu nu s-a oprit aici. A spus: «Elena, am creat Cystiolla nu ca să câștig bani. L-am creat după ce propria mea soră a trecut prin același iad ca dumneavoastră. Am văzut durerea ei. Am văzut umilința ei. Am văzut cum își cheltuia pensia pe medicamente care nu funcționau. Și am jurat: preparatul meu va fi accesibil fiecărei femei, nu doar celor cu bani.',
    100: 'De aceea pentru femeile care vor citi scrisoarea dumneavoastră am făcut prețul de 111 lei. E sub costul de producție. De 4 ori mai ieftin decât prețul obișnuit și de 27 de ori mai ieftin decât ceea ce cheltuiți acum pe medicamente inutile. Pierd bani la fiecare cutie. Dar nu vreau încă o Elena care să scrie o astfel de scrisoare. Nu vreau încă o femeie să moară doar pentru că nu și-a permis adevărul».',
    101: '111 lei.',
    102: 'Două cine la restaurant. Trei pachete de cafea bună. O vizită la frizerie. Sau — o viață fără cistită. O viață fără durere. O viață fără frică.',
    103: 'Nu știu cât mi-a mai rămas. Medicii spun — trei luni. Scriu această scrisoare și plâng. Nu pentru mine. Pentru voi. Pentru cei care acum beau Monural și gândesc:',
    104: '«Nimic grav, e doar cistită». Pentru cei care mănâncă ciocolată seara și nu știu că hrănesc bacteriile din vezica urinară. Pentru cei care se trezesc noaptea la toaletă și atribuie asta vârstei. Pentru cei care acum se uită la prețul de 111 lei și gândesc: «Probabil e o înșelătorie, nu poate fi atât de ieftin».',
    105: 'Poate. Pentru că doctorul Popescu nu e o companie farmaceutică. E medic. A văzut moartea. Și nu vrea să câștige pe ea.',
    106: 'Nu e vârsta. Nu e doar cistită. E rezistența la insulină. Și vă va ucide, dacă nu vă opriți.',
    107: 'Opriți-vă. Înainte să fie prea târziu. 111 lei. Asta e prețul vieții mele. Nu repetați greșeala mea.',
    108: 'Elena Matei, Brașov.<br>12 martie 2026.',
    110: 'P.S. de la redacție. Elena a încetat din viață pe 18 martie. Publicăm această scrisoare conform ultimei ei dorințe. Doctorul Irnel Popescu a confirmat: pentru cititoarele acestei scrisori prețul la Cystiolla este de 111 lei. Cantitatea de cutii la acest preț este limitată.',
    112: 'Apăsați butonul de mai jos pentru a lăsa o cerere. Un consultant vă va suna în 10 minute. Fără plată în avans. Plata doar la primire.',
    113: 'Lăsați numărul de telefon chiar acum. Până nu șterg scrisoarea Elenei — companiile farmaceutice deja cer să o scoată.',
}

# Bold segments within paragraphs (run index -> bold text in Romanian)
BOLD = {
    10: [('цистит', 'cistita'), ('резь', 'arsură'), ('Жжение', 'Înțepătură')],
    12: [('снова', 'din nou'), ('Ну что вы хотите, возраст, менопауза. Пейте клюквенный морс и смиритесь', 'Ei bine, ce vreți, vârsta, menopauza. Beți suc de merișoare și învățați să trăiți cu asta')],
    13: [('Я тратила на это больше 3.000 леев в год. За 10 лет — 30.000 леев', 'Cheltuiam pe toate acestea peste 3.000 de lei pe an. În 10 ani — 30.000 de lei'), (' Я отдала эти деньги за то, чтобы медленно умирать.', ' Am dat acești bani ca să mor încet.')],
    17: [('Постоянная усталость', 'Oboseală constantă')],
    18: [('Тяга к сладкому', 'Poftă de dulce')],
    19: [('Сухость во рту.', 'Uscăciune în gură.')],
    20: [('Кожный зуд', 'Mâncărime a pielii')],
    21: [('Ухудшение зрения', 'Agravarea vederii')],
    22: [('Медленно заживающие ранки', 'Răni care se vindecă greu')],
    23: [('Частые позывы в туалет ночью', 'Urgențe frecvente la toaletă noaptea')],
    25: [('Ни один не сказал: «Давайте проверим инсулин»', 'Niciunul nu a spus: «Hai să verificăm insulina»')],
    28: [('«А вы знаете, что у вас преддиабет?»', '«Știți că aveți prediabet?»')],
    31: [('в 4 раза выше нормы. ', 'de 4 ori mai mare decât normal. '), ('Но сахар всё равно просачивался в мочу.', 'Dar zahărul tot se scurgea în urină.'), ('И через неделю вырастала новая колония.', 'Și după o săptămână creștea o colonie nouă.')],
    33: [('«Почему никто не сказал мне этого раньше? Почему ни один уролог не проверил инсулин?»', '«De ce nimeni nu mi-a spus asta mai devreme? De ce niciun urolog nu a verificat insulina?»')],
    34: [('Уролог ищет бактерии в моче. Эндокринолог ищет проблемы с гормонами. А то, что между ними — ничья территория. Вы попали в мёртвую зону медицины', 'Urologul caută bacterii în urină. Endocrinologul caută probleme hormonale. Iar ce e între ele — teritoriu nimănui. Ați ajuns în zona moartă a medicinei')],
    36: [('пиелонефрит', 'pielonefrită'), ('почечную недостаточность', 'insuficiență renală cronică'), ('диабет 2 типа', 'diabet de tip 2 adevărat')],
    37: [('Сахар и инфекция убили мои почки.', 'Zahărul și infecția mi-au ucis rinichii.')],
    44: [('там сахар', 'acolo e zahăr')],
    46: [('разрушает сосуды', 'distruge vasele de sânge'), ('Слизистая истончается', 'Mucoasa se subțiază'), ('Бактерии проникают вглубь тканей', 'Bacteriile pătrund adânc în țesuturi')],
    47: [('Год за годом. Антибиотик за антибиотиком. Пока почки не скажут «хватит».', 'An după an. Antibiotic după antibiotic. Până când rinichii spun «destul».')],
    50: [('эндокринологию и урологию', 'endocrinologia și urologia')],
    52: [('Елена, ваш цистит никогда не был инфекцией. Он был симптомом инсулинорезистентности. Если бы вам 10 лет назад прописали берберин и альфа-липоевую кислоту, вы бы сейчас были здоровы', 'Elena, cistita dumneavoastră nu a fost niciodată o infecție. A fost un simptom al rezistenței la insulină. Dacă acum 10 ani vi s-ar fi prescris berberină și acid alfa-lipoic, acum ați fi sănătoasă')],
    59: [(' Я искал природные компоненты', ' Căutam componente naturale'), ('убрать сахар из мочи и восстановить чувствительность клеток к инсулину', 'să scoată zahărul din urină și să restabilească sensibilitatea celulelor la insulină'), ('это оно', 'asta e')],
    60: [('Три компонента', 'Trei componente'), ('на миллиграмм в сторону, и эффект как от плацебо.', 'cu un miligram în plus sau în minus, efectul e ca la placebo.')],
    61: [('Первый: берберин.', 'Primul: berberină.')],
    63: [('Второй: альфа-липоевая кислота.', 'Al doilea: acid alfa-lipoic.')],
    65: [('Третий: экстракт корицы цейлонской', 'Al treilea: extract de scorțișoară de Ceylon')],
    67: [('Я смешал эти три компонента в одной формуле', 'Am amestecat aceste trei componente într-o singură formulă'), ('Cystiolla', 'Cystiolla')],
    96: [('это цена одного курса', 'asta e prețul unui singur curs')],
    98: [('Я отдала 30.000 леев за то, чтобы умерет', 'Am dat 30.000 de lei ca să mor')],
    99: [('не для того, чтобы заработать', 'nu ca să câștig bani')],
    100: [('цену 111 леев', 'prețul de 111 lei')],
    101: [('111 леев', '111 lei')],
    105: [('Может.', 'Poate.'), ('не фармкомпания.', 'nu e o companie farmaceutică.')],
}

ITALIC = {2: True, 108: True, 113: True}
BOLD_ITALIC = {1: True, 2: True, 108: True}

PIC_PARAS = {1: 1, 11: 2, 29: 3, 30: 3, 38: 4, 39: 4, 51: 5, 56: 6, 62: 7, 64: 8, 66: 9, 77: 10, 95: 11, 111: 12}

LIST_PARAS = set(range(17, 24)) | set(range(80, 90))

HEADING_PARAS = {8, 15, 27, 41, 49, 70, 77, 93, 111}


def apply_bold(text, idx):
  if idx not in BOLD:
    return text
  for _, ro_bold in BOLD[idx]:
    if ro_bold and ro_bold in text:
      text = text.replace(ro_bold, f'<b>{ro_bold}</b>', 1)
  return text


def img_tag(num):
    src = IMAGE_FILES.get(num, f'pic{num}.jpg')
    return f'<p style="text-align: center;"><img src="{src}" style="max-width: 700px; width: 100%;" loading="lazy"></p>'


def build():
    items = parse_docx('text.docx')
    lines = []
    pic_inserted = set()

    for i, item in enumerate(items):
        if i not in RO and i not in PIC_PARAS and i not in {3, 7, 14, 26, 40, 48, 69, 76, 91, 92, 109, 114}:
            continue

        if i in PIC_PARAS:
            num = PIC_PARAS[i]
            if i == 1:
                text = RO.get(1, '')
                lines.append(f'<p><b><i>{text}</i></b></p>\n<br>')
            if num not in pic_inserted:
                lines.append(img_tag(num))
                pic_inserted.add(num)
            if i == 77:
                text = RO.get(77, '')
                lines.append(f'<h1><b>{text}</b></h1>\n<br>')
            elif i == 111:
                pass
            continue

        if i in {3, 7, 14, 26, 40, 48, 69, 76, 91, 92, 109, 114}:
            lines.append('<br>')
            continue

        text = RO[i]
        text = apply_bold(text, i)

        if i == 0:
            lines.append(f'<h1 class="mainZag" style="font-weight: normal; font-size: 43px !important;"><span style="font-weight: bold;">{text}</span></h1>\n<br>')
        elif i in HEADING_PARAS:
            lines.append(f'<h1><b>{text}</b></h1>\n<br>')
        elif i in LIST_PARAS:
            if i == 17:
                lines.append('<ul>')
            lines.append(f'<li>{text}</li>')
            if i == 23 or i == 89:
                lines.append('</ul>\n<br>')
        elif ITALIC.get(i) and BOLD_ITALIC.get(i):
            lines.append(f'<p><b><i>{text}</i></b></p>\n<br>')
        elif ITALIC.get(i):
            lines.append(f'<p><i>{text}</i></p>\n<br>')
        else:
            lines.append(f'<p>{text}</p>\n<br>')

    return '\n'.join(lines)


def parse_docx(path):
    with zipfile.ZipFile(path, 'r') as z:
        doc = z.read('word/document.xml')
    root = ET.fromstring(doc)
    body = root.find(f'{W}body')
    items = []
    for elem in body:
        tag = elem.tag.split('}')[-1]
        if tag == 'p':
            para = {'type': 'p', 'runs': []}
            pPr = elem.find(f'{W}pPr')
            if pPr is not None:
                pStyle = pPr.find(f'{W}pStyle')
                if pStyle is not None:
                    para['style'] = pStyle.get(f'{W}val')
                numPr = pPr.find(f'{W}numPr')
                if numPr is not None:
                    para['list'] = True
            items.append(para)
    return items


if __name__ == '__main__':
    html = build()
    with open('content_ro.html', 'w', encoding='utf-8') as f:
        f.write(html)
    print('Done:', len(html), 'chars')
