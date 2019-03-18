// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

$(function () {
    var $buttonEl = $('#launchPicker');
    var $iframeEl = $('#embed');

    function handlePreviewResponse(response) {
        let previewUrl = response.getUrl || '';

        $iframeEl[0].src = previewUrl;
    }

    function postData(url = ``, accessToken = '') {
        return fetch(url, {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "Authorization": `Bearer ${accessToken}`
            }
        })
        .then(response => response.json());
    }

    function getEmbedUrl(endpoint, driveId, itemId) {
        return `${endpoint}/drives/${driveId}/items/${itemId}/preview`;
    }

    function onSuccess(response) {
        console.log('success', response);

        let accessToken = response.accessToken;
        let apiEndpoint = response.apiEndpoint || 'https://graph.microsoft.com/v1.0/';
        let file = response.value[0] || {};
        let fileId = file.id || '';
        let parentDrive = file.parentReference || {};
        let driveId = parentDrive.driveId || '';

        let postUrl = getEmbedUrl(apiEndpoint, driveId, fileId);

        postData(postUrl, accessToken)
            .then(handlePreviewResponse)
            .catch(onError);
    }

    function onCancel() {
        console.log('Action canceled');
    }

    function onError(err) {
        console.error(err);
    }

    function launchPicker() {
        var options = {
            clientId: 'bb11e7d9-afdb-4c20-b282-80b9c60a7bd1',
            accountSwitchEnabled: false,
            success: onSuccess,
            cancel: onCancel,
            error: onError
        };

        OneDrive.open(options);
    }

    $buttonEl.on('click', launchPicker);
});
