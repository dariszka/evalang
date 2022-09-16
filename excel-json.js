const xlsx = require("xlsx")
const fs = require("fs")

const wb = xlsx.readFile("./Data.xlsx")

const ws = wb.Sheets["Sheet1"]
const data = xlsx.utils.sheet_to_json(ws);
const matches = []

let fullListOfResults = []

match()

function match() {
    let names = data.map(x => x.name);
    let numbers = data.map(x => x.numbers);
    let spokenLs = data.map(x => x.spokenL);
    let desiredLs = data.map(x => x.desiredL);

    for (let i = 0; i < data.length; i++) {

        let firstLoopPerson = names[i]
        let firstLoopPersonsNumber = numbers[i]
        let firstLoopSpokenL = spokenLs[i]
        let firstLoopDesiredL = desiredLs[i]

            for (let i = 0; i < spokenLs.length; i++) {
                let secondLoopPerson = names[i]
                let secondLoopPersonsNumber = numbers[i]
                let secondLoopSpokenL = spokenLs[i]
                let secondLoopDesiredL = desiredLs[i]

                if (firstLoopSpokenL == secondLoopDesiredL) {
                    matches[0] = firstLoopPerson

                    if(firstLoopDesiredL == secondLoopSpokenL) {
                        matches[1] = secondLoopPerson

                        fullListOfResults[i] = [matches[0], matches[1]]
                    }
                }   
            }
    } readFullListOfResults()
}
    

function readFullListOfResults() {
    fullListOfResults.forEach((pair) => { 	
        pair.sort()
    });

    let alphabeticPairsListOfResults = fullListOfResults.filter(function(element) {
        return element !== undefined;
    });

    for (let i = 0; i < alphabeticPairsListOfResults.length; i++) {
        let stringPair = alphabeticPairsListOfResults[i].toString();
        let formattedStringPair = stringPair.replace(","," and ");
        alphabeticPairsListOfResults[i] = formattedStringPair
    }
    let finalFullMatches =  [...new Set(alphabeticPairsListOfResults)];
    displayResults(finalFullMatches)
}

function displayResults(fullMatches) {
    let formattedFullMatches = 'Full matches: \n' + fullMatches.toString().replace(/,/g,'\n')
    
    let textResult =  formattedFullMatches
    console.log(textResult)
}
