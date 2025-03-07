(function (window, undefined) {
    window.Asc.plugin.init = function () {
        main()
    };

    window.Asc.plugin.button = function (id) {
        this.executeCommand("close", "");
    };

    function F2Enter() {
        const activeSheet = Api.GetActiveSheet()
        const selectedRange = Api.GetSelection()
        const selectedAddress = selectedRange.Address
        if (/^\$[A-Z]+:\$[A-Z]+$/.test(selectedAddress) || selectedAddress.endsWith('1048576')) { // если выделили столбец полностью
            console.log(Common.UI.alert({
                title: 'Ошибка',
                msg: `Выбран целый столбец: ${selectedRange.Address}`
            }))
        } else {
            if (selectedRange.areas.length > 1) { //если выделели несколько диапазонов через Ctrl
                let ranges = []
                selectedRange.areas.forEach(area => { //находим адреса этих диапазонов
                    let range = []
                    let startRow = area.kb.r1
                    let startColumn = area.kb.ia
                    let endRow = area.kb.r2
                    let endColumn = area.kb.ra
                    range.push(startRow, startColumn, endRow, endColumn)
                    // console.log(range)
                    ranges.push(range)
                });

                ranges.forEach(range => {
                    for (let i = range[0]; i <= range[2]; i++) {
                        for (let j = range[1]; j <= range[3]; j++) {
                            let cell = activeSheet.GetRangeByNumber(i, j)
                            transformCellsVisability(cell)
                        }
                    }

                })
            } else { // если один диапазон
                selectedRange.ForEach(cell => {
                    transformCellsVisability(cell)
                })
            }
        }

        function transformCellsVisability(cell) {
            let cellValue = cell.GetValue();
            if (cellValue.length > 0) {
                if (cellValue.startsWith('=')) {
                    cell.SetNumberFormat('General');
                    cell.SetValue(cellValue);
                } else if (Number(cellValue)) {
                    cell.SetNumberFormat('General');
                    cell.SetValue(cellValue);
                }
            }
        }

    }

    async function main() {
        return new Promise((resolve) => {
            window.Asc.plugin.callCommand(F2Enter, true, true, function (value) {
                resolve(value);
            })
        })
    }

})(window, undefined);