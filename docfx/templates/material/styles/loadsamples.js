/**
 * This file is unique for each sample browser. It contains the logic specific to each repo for loading the samples as needed.
 */
var jsonPath = document.location.origin +"/samples.json";

/**
 * Reads a sample metadata and returns a pre-populated HTML element
 * @param {*} sample 
 * @returns 
 */
function loadSample(sample, filter) {
    try {
        // _ is missing
        // var title = _.escape(sample.title);
        // var escapedDescription = _.escape(sample.shortDescription);

        var title = sample.title;
        var escapedDescription = sample.shortDescription;

        var shortDescription = sample.shortDescription; //.length > 80 ? sample.shortDescription.substr(0, 77)  : sample.shortDescription;
        var thumbnail = document.location.origin + "/assets/nopreview.png";
        

        if (sample.thumbnails && sample.thumbnails.length > 0) {
          thumbnail = sample.thumbnails[0].url;
        }

        var libraries = "";
        var operations = "";
        var products = "";


        var metadata = sample.metadata;
        metadata.forEach(meta => {
          if (libraries !== "") {
            libraries = libraries + ", ";
          }
          libraries = libraries + meta.key.toLowerCase();
        });
        
        // Operations are represented as catagories in the sample.json file
        sample.categories.forEach(cata => {
          if (operations !== "") {
            operations = operations + ", ";
          }
          operations = operations + cata.toLowerCase();
        });

        // Products
        sample.products.forEach(product => {
          if (products !== "") {
            products = products + ", ";
          }
          products = products + product.toLowerCase();
        });

        var modified = new Date(sample.updateDateTime).toString().substr(4).substr(0, 12);
        var authors = sample.authors;
        var authorsList = "";
        var authorAvatars = "";
        var authorName = "";
        var authorsGitHub = "";
        var productTag = sample.products[0].toLowerCase();
        var productName = sample.products[0];


        // Build the authors array
        if (authors.length < 1) {
          console.log("Sample has no authors", sample);
        } else {
          authors.forEach(author => {
            if (authorsList !== "") {
              authorsList = authorsList + ", ";
            }
            authorsList = authorsList + author.name;
            authorsGitHub = authorsGitHub + " " + author.gitHubAccount;

            var authorAvatar = `<div class="author-avatar">
              <div role="presentation" class="author-coin">
                <div role="presentation" class="author-imagearea">
                  <div class="image-400">
                    <img class="author-image" loading="lazy" src="${author.pictureUrl}" alt="${author.name}" title="${author.name}">
                  </div>
                </div>
              </div>
            </div>`;
            authorAvatars = authorAvatar + authorAvatars;
          });

          authorName = authors[0].name;
          if (authors.length > 1) {
            authorName = authorName + ` +${authors.length - 1}`;
          }
        }

        // Extract tags
        var tags = "";
        $.each(sample.tags, function (_u, tag) {
          tags = tags + "#" + tag + ",";
        });

        // Build a keyword tag for searching
        var keywords = title + " " + escapedDescription + " " + authorsList + " " + authorsGitHub + " " + tags;
        keywords = keywords.toLowerCase();

        // Build the HTML to insert
        var $items = $(`
<a class="sample-thumbnail" href="${sample.url}" data-modified="${sample.modified}" data-title="${title}" data-keywords="${keywords}" data-tags="${tags}" data-libraries="${libraries}" data-operation="${operations}" data-products="${products}">
  <div class="sample-inner">
    <div class="sample-preview">
      <img src="${thumbnail}" loading="lazy" alt="${title}">
    </div>
    <div class="sample-details">
      <div class="producttype-item ${productTag}">${productName}</div>
      <p class="sample-title" title="${sample.title}">${sample.title}</p>
      <p class="sample-description" title='${escapedDescription}'>${shortDescription}</p>
      <div class="sample-activity">
        ${authorAvatars}
        <div class="activity-details">
          <span class="sample-author" title="${authorsList}">${authorName}</span>
          <span class="sample-date">Modified ${modified}</span>
        </div>
      </div>
    </div>
  </div>
</a>`);

       return $items;
      } catch (error) {
        console.log("Error with one sample", error, sample);
      }
      return null;
}