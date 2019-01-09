const Scraper = require("image-scraper")
const fs = require('fs')

function downloadImages (category) {
  let scraper = new Scraper('https://rocket-league.com/items/' + category)
  let jsonFile = 'images/' + category + '.json'
  let images = []

  scraper.scrape(function (image) {
    images.push({
      name: image.attributes.alt,
      url: image.address,
    })
  })

  setTimeout(() => {
    fs.writeFile(jsonFile, JSON.stringify(images), (err) => {
      if (err) {
        console.error(err)
        return
      }

      console.log('Wrote ' + images.length + ' images to ' + jsonFile)
    })
  }, 5000)
}

const categories = [
  'bodies',
  'decals',
  'wheels',
  'boosts',
  'antennas',
  'toppers',
  'trails',
  'explosions',
  'paints',
  'banners',
  'crates',
]

categories.forEach(cat => {
  downloadImages(cat)
})