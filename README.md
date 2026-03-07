# CopyToGo - PowerPoint Add-in

![CopyToGo Logo](assets/logo-filled.png)

**CopyToGo** is een PowerPoint Add-in waarmee je eenvoudig eigenschappen van afbeeldingen kunt kopiëren en plakken. Ideaal voor het snel uitlijnen van screenshots en afbeeldingen in presentaties.

## 🎯 Functionaliteit

- **Kopiëren**: Sla positie, afmetingen en andere eigenschappen van een afbeelding op
- **Plakken**: Pas de opgeslagen eigenschappen toe op een andere afbeelding
- **Werkt met**: Screenshots, afbeeldingen, en andere PowerPoint-objecten

## 🚀 Installatie

### Stap 1: Download het manifest
Download het [`manifest.xml`](manifest.xml) bestand van deze repository.

### Stap 2: Installeer in PowerPoint
1. Open PowerPoint (Desktop of Online)
2. Ga naar **Invoegen** → **Mijn Add-ins**
3. Klik op **Bestand uploaden**
4. Selecteer het gedownloade `manifest.xml` bestand
5. Klik op **OK**

### Stap 3: Gebruik de add-in
1. De **CopyToGo** knop verschijnt in het lint (Home-tabblad)
2. Klik erop om het paneel te openen
3. Selecteer een afbeelding en klik op **Kopiëren**
4. Selecteer een andere afbeelding en klik op **Plakken**

## 💻 Ondersteunde Platforms

- ✅ PowerPoint Desktop (Windows)
- ✅ PowerPoint Desktop (Mac)
- ✅ PowerPoint Online (Browser)

## 🛠️ Lokale Ontwikkeling

Wil je de add-in zelf aanpassen of verder ontwikkelen?

### Vereisten
- Node.js (versie 14 of hoger)
- npm

### Installatie
```bash
# Clone de repository
git clone https://github.com/jouwgebruikersnaam/copytogo.git

# Ga naar de projectmap
cd copytogo

# Installeer dependencies
npm install

# Start de development server
npm start
