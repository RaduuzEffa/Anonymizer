const text = "Tady je daniela a oheň a 25.2.1979, uvidime jestli to funguje.";
const customWord = "daniela, oheň, 25.2.1979";
let foundEntities = [];
let entityIdCounter = 0;

const escapeRegExp = (string) => string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const words = customWord.split(',').map(w => w.trim()).filter(w => w.length > 0);

words.forEach(word => {
    const customRegex = new RegExp(escapeRegExp(word), 'gi');
    const m = text.match(customRegex);
    if (m) m.forEach(x => {
        foundEntities.push({ id: ++entityIdCounter, originalText: x, type: 'CUSTOM' });
    });
});

console.log(foundEntities);
