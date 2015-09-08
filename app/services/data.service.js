(function () {
  'use strict';

  angular.module('officeAddin')
    .service('dataService', ['sharePointUrl', '$http', '$q', dataService]);

  /**
   * Custom Angular service.
   */
  function dataService(sharePointUrl, $http, $q) {
    
    // public signature of the service
    return {
      getDocuments: getDocuments
    };

    /** *********************************************************** */

    function getValueFromResults(key, results) {
      var value = '';

      if (results !== null &&
        results.length > 0 &&
        key !== null) {
        for (var i = 0; i < results.length; i++) {
          var resultItem = results[i];

          if (resultItem.Key === key) {
            value = resultItem.Value;
            break;
          }
        }
      }

      return value;
    }
    
    function getDocuments(query) {
      var deferred = $q.defer();

      var searchQuery = "?querytext='" + query + " isdocument:1'&SelectProperties='HitHighlightedSummary,LastModifiedTime,Path,SPWebUrl,ServerRedirectedURL,SiteTitle,Title'&RowLimit=5&StartRow=0";

      $http({
        url: sharePointUrl + '/_api/search/query' + searchQuery,
        method: 'GET',
        headers: {
          'Accept': 'application/json;odata=nometadata'
        }
      }).success(function (data) {
        var documents = [];

        if (data.PrimaryQueryResult !== null) {
          data.PrimaryQueryResult.RelevantResults.Table.Rows.forEach(function (row) {
            var cells = row.Cells;

            var url = getValueFromResults('ServerRedirectedURL', cells);
            if (url === null) {
              url = getValueFromResults('Path', cells);
            }

            documents.push({
              url: url,
              title: getValueFromResults('Title', cells),
              summary: getValueFromResults('HitHighlightedSummary', cells).replace(/<(\/)?c\d>/g, '<$1mark>').replace(/<ddd\/>/g, ''),
              siteUrl: getValueFromResults('SPWebUrl', cells),
              siteTitle: getValueFromResults('SiteTitle', cells)
            });
          });
        }

        deferred.resolve(documents);
      }).error(function (err) {
        deferred.reject(err);
      });

      return deferred.promise;
    }
    
  }
})();