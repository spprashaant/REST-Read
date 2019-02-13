'use strict';

ExecuteOrDelayUntilScriptLoaded(initializePage, "sp.js");

function initializePage() {
    //#region Common Code
    function readAjaxCall(restUrl) {
        var dfd = $.ajax({
            url: restUrl,
            type: "GET",
            contentType: "application/json;odata=verbose",
            headers: {
                "accept": "application/json;odata=verbose"
            }
        });
        return dfd.promise();
    }
    function displayError(error) {
        $('#message').append("<br/> There was an error getting the data..");
        $('#message').append(error);
    }
    //#endregion

    // This code runs when the DOM is ready and creates a context object which is needed to use the SharePoint object model
    $(document).ready(function () {
        $("#readLists").click(readLists);
        $("#readListItems").click(readListItems);
        $("#readItemsPaged").click(readListItemsPaged);
        
    });

    //#region Lists
    function readLists() {
        var msg = "Getting all lists...";
        $('#message').text(msg);
        var restUrl = SP.Utilities.UrlBuilder.urlCombine(_spPageContextInfo.webServerRelativeUrl,
            "_api/web/lists");
        var readPromise = readAjaxCall(restUrl);
        readPromise.then(
            function (data, status, jqXHR) {
                displayLists(data);
            },
            function (jqXHR, status, error) {
                displayError(error);
            }
        );
    }
    function displayLists(data) {
        var msg = "<br/><br/><b>Lists on this site are:</b><ul>";
        if (data.d.results.length > 0) {
            $.each(data.d.results, function (index, value) {
                msg += "<li>" + value.Title + "</li>";
            });
        }
        else {
            msg += "<li>" + data.d.Title + "</li>";
        }
        msg += "</ul>";
        $('#message').append(msg);
    }
    //#endregion

    //#region List Items
    function readListItems() {
        var msg = "Getting all lists items...";
        var listName = "SampleData";
        $('#message').text(msg);
        var restUrl = SP.Utilities.UrlBuilder.urlCombine(_spPageContextInfo.webServerRelativeUrl,
            "_api/web/lists/GetByTitle('" + listName + "')/items");
        var readPromise = readAjaxCall(restUrl);
        readPromise.then(
            function (data, status, jqXHR) {
                displayListItems(data);
            },
            function (jqXHR, status, error) {
                displayError(error);
            }
        );
    }

    function displayListItems(data) {
        var msg = "<br/><br/><b>List Items are:</b><ul>";
        if (data.d.results.length) {
            if (data.d.results.length > 1) {
                msg += data.d.results.length + " items:<ul>";
                $.each(data.d.results, function (index, value) {
                    msg += "<li>" + value.Title + "</li>";
                });

            }
            else {
                msg += "1 item:<ul>";
                msg += "<li>" + data.d.results[0].Title + "</li>";
            }
            msg += "</ul>";
        }
        else {
            msg += "no items";
        }

        if (data.d.__next) {
            msg += "<br />";
            msg += "Results are paged - next set of results available here: <br />" + data.d.__next;
        }
        $('#message').append(msg);
    }
    //#endregion

    //#region Paging List Items
    function readListItemsPaged() {
        var msg = "Getting next page lists items...";
        var listName = "SampleData";
        var pageNum = 0;
        var pageSize = 3;
        $('#message').text(msg);
        var restUrl = SP.Utilities.UrlBuilder.urlCombine(_spPageContextInfo.webServerRelativeUrl,
            "_api/web/lists/GetByTitle('" + listName + "')/items?$top=" + pageSize);
        //need to add skiptoken unencoded or else the % in %26 and %3d gets encoded to %25 which breaks things
        restUrl = encodeURI(restUrl) + "&$skiptoken=Paged%3dTRUE%26p_ID%3d" + pageNum * pageSize;

        var readPromise = readAjaxCall(restUrl);
        readPromise.then(
            function (data, status, jqXHR) {
                displayListItems(data);
            },
            function (jqXHR, status, error) {
                displayError(error);
            }
        );
    }

    function displayListItemsPaged(data) {
        var msg = "<br/><br/><b>List Items are:</b><ul>";
        if (data.d.results.length) {
            if (data.d.results.length > 1) {
                msg += data.d.results.length + " items:<ul>";
                $.each(data.d.results, function (index, value) {
                    msg += "<li>" + value.Title + "</li>";
                });

            }
            else {
                msg += "1 item:<ul>";
                msg += "<li>" + data.d.results[0].Title + "</li>";
            }
            msg += "</ul>";
        }
        else {
            msg += "no items";
        }

        if (data.d.__next) {
            msg += "<br />";
            msg += "Results are paged - next set of results available here: <br />" + data.d.__next;
        }
        $('#message').append(msg);
    }
    //#endregion
}
