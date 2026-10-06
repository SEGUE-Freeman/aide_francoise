# Codification des emplacements de stock – CAMEG Atakpamé

Squelette d'un système pour que **chaque produit ait une adresse** dans le magasin
et que n'importe qui (stagiaire, remplaçant, logisticien) puisse le retrouver.

| Fichier | Contenu |
|---|---|
| [`Modele_Codification_Stock_CAMEG.xlsx`](Modele_Codification_Stock_CAMEG.xlsx) | Modèle Excel prêt à remplir (Zones, Emplacements, Produits, Stock, Recherche, Étiquettes) |
| [`outils/generer_modele.py`](outils/generer_modele.py) | Script qui régénère le modèle (`python outils/generer_modele.py`) |

---

## 1. Existe-t-il une norme ?

**Il n'y a pas de norme internationale qui impose un format de code d'emplacement.**
Ce qu'on trouve partout (entrepôts, logiciels de gestion d'entrepôt / WMS, centrales
d'achat de médicaments) est une **convention** :

```
ZONE  –  ALLÉE / ÉTAGÈRE  –  NIVEAU  –  POSITION
```

Les vraies références à citer dans un rapport de stage :

| Référence | Ce qu'on y trouve |
|---|---|
| **OMS – Good storage and distribution practices for medical products** (WHO Technical Report Series n° 1025, annexe 7, 2020) | Exigences de stockage : zones séparées (quarantaine, rejetés, produits sous contrôle), conditions de température, traçabilité des lots, FEFO |
| **USAID / JSI – Guidelines for the Storage of Essential Medicines and Other Health Commodities** (2003) | Guide pratique très concret : organisation du magasin, rangement, étiquetage, FEFO, sécurité, chaîne de froid |
| **USAID DELIVER PROJECT – The Logistics Handbook** (2011) | Gestion logistique des produits de santé dans les pays à ressources limitées (inventaire, fiches de stock, SIGL) |
| **MSH – MDS-3 : Managing Access to Medicines and Health Technologies** (2012) | Gestion des magasins pharmaceutiques, classifications ABC et VEN |
| **GS1** (GLN, codes-barres GS1 DataMatrix des médicaments) | Pour plus tard : identification normalisée des sites et lecture par scanner |

Ces documents sont gratuits en ligne (taper leur titre exact dans un moteur de recherche).

## 2. Le squelette : 3 tables reliées

```
 EMPLACEMENTS                PRODUITS                    STOCK (où est chaque lot)
 ────────────                ────────                    ─────────────────────────
 Code emplacement ◄──┐       Code article ◄──────┐       Emplacement   ──► Emplacements
 Zone                │       DCI, dosage, forme  │       Code article  ──► Produits
 Étagère             │       Zone requise        │       N° de lot
 Niveau              └────── Emplacement principal       Date de péremption
 Position                    Classe ABC / VEN            Quantité
```

Idée clé : **le code produit ne contient pas l'emplacement**. Si on déplace un produit,
on change une ligne dans *Stock*, pas son code. Un même produit peut avoir plusieurs
lots à plusieurs endroits (casier de prélèvement + réserve sur palette + quarantaine).

## 3. Le format de code proposé

```
A03-2-04
│ │  │ └─ Position : 4e case en partant de la gauche (01, 02, …)
│ │  └─── Niveau   : 2e niveau en partant du bas (1 = sol)
│ └────── Étagère  : n° 03 dans la zone (01 à 99)
└──────── Zone     : A = stock général
```

| Zone | Usage |
|---|---|
| **A** | Stock général, température ambiante |
| **F** | Chaîne de froid (+2 °C / +8 °C) |
| **S** | Sécurisé (stupéfiants, psychotropes, produits de valeur) |
| **N** | Inflammables (alcool, éther…) |
| **P** | Palettes / vrac (perfusions, cartons de SRO…) |
| **Q** | Quarantaine (périmés, abîmés, retours, rappels de lots) |
| **R / E** | Réception / Expédition |

Règles :

1. **Un code = un seul endroit physique**, jamais réutilisé pour un autre endroit.
2. **Longueur fixe** (`01`, pas `1`) : le tri alphabétique suit alors l'ordre de marche dans le magasin.
3. **Pas de lettres I et O** (confusion avec 1 et 0).
4. **Garder les codes déjà peints** : les étagères sont déjà codifiées. Le modèle a une
   colonne « Ancien code » pour faire la correspondance. Si les codes existants sont
   cohérents, on peut même les garder tels quels comme partie « étagère ».
5. **Étiqueter chaque case** (onglet *Étiquettes* du modèle). Un code qui n'est pas affiché
   sur l'étagère ne sert à rien.
6. Si plusieurs magasins/agences doivent un jour partager le système, ajouter un préfixe de site
   devant (ex. `ATK-A03-2-04`), à choisir avec la direction.

## 4. Où ranger quoi (logique de placement)

1. **D'abord les conditions de stockage** → la zone (froid, sécurisé, inflammable, quarantaine…).
2. **Ensuite, dans la zone**, un seul critère simple et constant, par exemple :
   forme galénique (comprimés / injectables / sirops…) puis **ordre alphabétique de la DCI**.
   (Le rangement par classe thérapeutique est aussi possible ; l'essentiel est d'en choisir un et de s'y tenir.)
3. **Produits à forte rotation (classe A)** près de la sortie / zone de préparation, à hauteur de main.
4. **Lourds et volumineux en bas**, jamais à même le sol (palettes).
5. **FEFO** : le lot qui périme le plus tôt devant / en haut de la pile.

## 5. Démarche sur le terrain (≈ 2 à 4 semaines à temps partiel)

1. **Présenter l'idée au tuteur et au logisticien avant de toucher aux étagères** (même si ce
   n'est pas demandé, c'est leur magasin : leur accord est indispensable et ça valorise le stage).
2. **Faire le plan du magasin** à la main : zones, étagères, sens de circulation.
3. **Relever toutes les étagères** (nombre de niveaux, de cases) → onglet *Emplacements*.
   Les codes se calculent tout seuls.
4. **Lister les produits** avec les codes articles officiels CAMEG → onglet *Produits*.
5. **Inventaire physique** : pour chaque case, noter produit, lot, péremption, quantité → onglet *Stock*.
6. **Imprimer et coller les étiquettes**.
7. **Tester** : demander à un collègue de retrouver 10 produits avec l'onglet *Recherche*.
8. **Fixer la règle de mise à jour** : tout rangement ou déplacement est noté le jour même
   (sinon le fichier devient faux en quelques semaines).

## 6. Quelle plateforme / quelle IA ?

**Pour l'outil :**

| Étape | Outil | Pourquoi |
|---|---|---|
| Maintenant | **Excel** (ou Google Sheets) avec le modèle fourni | Gratuit, déjà installé, aucun accord informatique nécessaire |
| Ensuite | **QR codes / codes-barres** sur les étiquettes | Lecture au téléphone ou au scanner |
| Plus tard | Vérifier d'abord le **logiciel déjà utilisé par la CAMEG** (beaucoup de logiciels de gestion de stock ont un champ « emplacement » inutilisé). Sinon : **Odoo Inventaire** (édition communautaire gratuite, emplacements à plusieurs niveaux), **Open mSupply** (open source, conçu pour les produits de santé) | Décision de la direction / du service informatique, pas d'une stagiaire seule |

**Pour l'IA :** Claude, ChatGPT ou Gemini font très bien l'affaire pour structurer le système,
écrire des formules Excel ou rédiger une procédure. Pour trouver des documents sourcés
(OMS, USAID), **Perplexity** est pratique car il cite ses sources.

Conseils :

- **Ne jamais coller de données internes de la CAMEG** (quantités, prix, fournisseurs) dans une IA publique.
- **Toujours vérifier** une « norme » citée par une IA dans le document original : les IA inventent parfois des références.
- Exemple de question efficace :
  > « Je suis stagiaire dans un magasin régional de médicaments essentiels au Togo. J'ai un
  > fichier Excel avec les onglets Emplacements, Produits et Stock. Écris-moi une procédure
  > d'une page, en français simple, pour mettre à jour l'emplacement d'un produit quand on le déplace. »

## 7. Régénérer le modèle

```bash
python outils/generer_modele.py Modele_Codification_Stock_CAMEG.xlsx
```

Nécessite Python 3 et `openpyxl`. Le fichier est compatible Excel 2010 et plus, LibreOffice et
Google Sheets (formules classiques INDEX/MATCH, pas de fonctions récentes).
