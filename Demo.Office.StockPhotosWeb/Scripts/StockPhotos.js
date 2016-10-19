
var appId = "[App Identifier]";
var resultsHeaderString = "<p>Found {0} results.</p>";
$.support.cors = true;

function searchPhotos(term, page) {
    page = ((page == null || page == 'undefined' || page == '') ? 1 : page);
    term = ((term == null || term == 'undefined' || term == '') ? $("#searchTerm").val() : term);
    var baseUrl = "https://api.unsplash.com/search/photos?client_id=" + appId + "&page=" + page + "&query=" + term;
    $.ajax({
        url: baseUrl,
        type: "GET",
        cache: false,
        crossDomain: true,
        dataType: "json",
        contentType: "application/json;charset=utf-8"
    }).done(function (data) {
        try {
            var _count = data.total;
            var _pages = data.total_pages;
            var _results = data.results;

            //Header
            $("#photoSearchResults_Header").html(resultsHeaderString.replace("{0}", _count));

            //Body
            var bodyHtml = "";
            $.each(_results, function (i, item) {
                var imgHtml = 
                bodyHtml += "<div class='item' data-w='150' data-h='150'><img data-src='" + item.urls.small + "' src='/images/blank.gif' height='150' width='150' onclick='insertImage(\"" + item.urls.full + "\");return false;'></div>";
            });
            $("#photoSearchResults_Items").html(bodyHtml);
            $(".flex-images").flexImages({ rowHeight: 150});

            //Footer
            var _baseFooterItem = "&nbsp;&nbsp;{0}&nbsp;&nbsp;"
            var _footerHtml = "<p>Page"
            var _resultPages = [];
            if (parseInt(_pages) > 1) {
                for (var i = 1; i <= _pages; i++) {
                    var curString = i.toString();
                    if (curString === page) {
                        _footerHtml += _baseFooterItem.replace("{0}", curString);
                    } else {
                        _footerHtml += _baseFooterItem.replace("{0}", "<a href='' onclick='searchPhotos(\"" + term + "\",\"" + curString + "\");return false;'>" + curString + "</a>");
                    }
                }
            } else {
                _footerHtml += _baseFooterItem.replace("{0}", page);
            }
            _footerHtml += "</p>";
            $("#photoSearchResults_Footer").html(_footerHtml);
        } catch (err) {
            alert(err);
        }
    }).fail(function (req, msg) {
        $("#photoSearchResults").html("<p>Request failed: " + msg + "</p>");
    });
}

function insertImage(url) {
    var imgHtml = "<img src='" + url + "' + alt='' />";
    Office.context.document.setSelectedDataAsync(
        imgHtml, { coercionType: "html" },
        function (asyncResult) {
            if (asyncResult.status == "failed") {
                write('Error: ' + asyncResult.error.message);
        }
    });
}