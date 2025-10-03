const fs = require('fs');
const path = require('path');


function readJsonFile(fileName) {
    let content = fs.readFileSync(path.join(__dirname, fileName), 'utf8');
    if (content.charCodeAt(0) === 0xFEFF) {
        content = content.slice(1);
    }
    return JSON.parse(content);
}

class Converter {
    constructor() {
        this.ar = readJsonFile('ar.json');
    }

    // Character reordering logic (equivalent to ci)
    ci(w) {
        let cY = 0;
        w = w.split(''); // Convert string to array for easier manipulation
        for (let i = 0; i < w.length; i++) {
            // Pre-kar movement
            if (i < w.length && this.ao(w[i])) {
                let j = 1;
                while (i - j >= 0 && this.v(w[i - j])) {
                    if (i - j <= cY) break;
                    if (i - j - 1 >= 0 && this.D(w[i - j - 1])) j += 2;
                    else break;
                }
                let temp = [...w.slice(0, i - j), w[i], ...w.slice(i - j, i), ...w.slice(i + 1)];
                w = temp;
                cY = i + 1;
                i = cY;
                continue;
            }

            // RA + HALANT + vowel handling
            if (i < w.length - 1 && this.D(w[i]) && w[i - 1] === 'র' && !this.D(w[i - 2])) {
                let j = 1;
                let aZ = 0;
                while (true) {
                    if (i + j >= w.length) break;
                    if (this.v(w[i + j]) && this.D(w[i + j + 1])) j += 2;
                    else if (this.v(w[i + j]) && this.ao(w[i + j + 1])) {
                        aZ = 1;
                        break;
                    } else break;
                }
                const temp = [
                    ...w.slice(0, i - 1),
                    ...w.slice(i + j + 1, i + j + aZ + 1),
                    ...w.slice(i + 1, i + j + 1),
                    w[i - 1],
                    w[i],
                    ...w.slice(i + j + aZ + 1)
                ];
                w = temp;
                i += (j + aZ);
                cY = i + 1;
                continue;
            }
        }
        return w.join('');
    }

    // Conversion function
    Convert(line) {
        //let az = "bangla";
        let G  = this.ar;

        // Fix combined vowels
        line = line.replace(/ো/g, "ো");
        line = line.replace(/ৌ/g, "ৌ");

        //Fix nukta mixed characters...
        line = line.replace(/য়/g, "q");
        line = line.replace(/ঢ়/g, "p");
        line = line.replace(/ড়/g, "o");

        // Reorder vowels/consonants
        line = this.ci(line);

        // Apply mapping
        for (const key in G) {
            const re = new RegExp(key, "g");
            line = line.replace(re, G[key]);
        }
        

        return line;
    }

    // Helper functions
    bA(e) { return '০১২৩৪৫৬৭৮৯'.includes(e); }
    ao(e) { return 'িৈে'.includes(e); }  // Pre-kar
    aJ(e) { return 'াোৌুূীৃ'.includes(e); } // Post-kar
    ah(e) { return this.ao(e) || this.aJ(e); }
    v(e) { return 'কখগঘঙচছজঝঞটঠডঢণতথদধনপফবভমশষসহযরলয়ংঃঁৎ'.includes(e); }
    D(e) { return e === '্'; }
}

module.exports = Converter;
