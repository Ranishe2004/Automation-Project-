# Cockpit de valorisation - 5 methodes

Interface web de valorisation d'entreprises cotees. Moteur de calcul live,
graphiques interactifs, sliders de sensibilite, export Excel / PDF.
Donnees par defaut : HPS (Hightech Payment Systems), Bourse de Casablanca.

Methodes : DCF, Comparables boursiers, Patrimoniale (ANR), Dividendes (DDM), EVA/MVA.

## Lancement

Prerequis : Node.js 18 ou superieur.

    npm install
    npm run dev

Ouvre automatiquement http://localhost:5173

## Production

    npm run build      # genere /dist
    npm run preview    # sert le build localement

## Stack

- React 18 + Vite
- Recharts (waterfalls, barres, courbes)
- SheetJS / xlsx (export Excel)
- CSS scope maison (aucune dependance Tailwind)

## Structure

    src/App.jsx      tout le cockpit (moteur + UI + styles)
    src/main.jsx     point d'entree React
    src/index.css    reset minimal

## Exports

- Excel : bouton "Excel" en haut a droite (6 onglets : Synthese + une feuille par methode).
- PDF : bouton "PDF" (boite d'impression du navigateur, feuille de style claire dediee).

## Conventions de calcul

- Multiples EV (EV/EBITDA, EV/EBIT, EV/CA) nets de la dette nette pour le passage
  a la valeur des fonds propres.
- Actualisation standard des EVA au taux WACC (PV).
- Tous les parametres sont editables : les valeurs par defaut peuvent etre remplacees
  pour toute autre societe.
