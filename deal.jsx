try {
    var doc = app.activeDocument;
    doc.suspendHistory("Rename Layers", "processLayers()");

    function processLayers() {
        var selectedLayers = getSelectedLayers();
        
        if (selectedLayers.length === 0) {
            alert("请先选择要处理的图层");
            return;
        }

        for (var i = 0; i < selectedLayers.length; i++) {
            var layer = selectedLayers[i];
            alert(layer.name);
        }
    }

    // 获取所有选中图层（支持多选）
    function getSelectedLayers() {
        var ActLay = app.activeDocument.activeLayer;
        ActLay.allLocked = true;
        var selLayers = new Array();
        
        // 遍历所有顶层图层
        traverseLayers(app.activeDocument.layers, selLayers);
        
        ActLay.allLocked = false;
        return selLayers;
    }

    // 递归遍历图层及图层组
    function traverseLayers(layers, selLayers) {
        for (var i = 0; i < layers.length; i++) {
            var layer = layers[i];

            // 如果是图层组（LayerSet），则递归处理
            if (layer.typename == "LayerSet") {
                traverseLayers(layer.layers, selLayers);
            }
            
            if (layer.allLocked == true) {
                selLayers.push(layer);
            }
        }
    }

} catch (e) {
    alert("错误: " + e.message + "\n行号: " + e.line);
}
