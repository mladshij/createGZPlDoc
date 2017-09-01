package main

import (
	"fmt"
	"io/ioutil"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/extrame/xls"
	"github.com/tealeg/xlsx"
	"golang.org/x/text/encoding/charmap"
	"golang.org/x/text/transform"
)

const rtLive = 1
const rtOffice = 2

type roomID struct {
	Number int
	Type   int
}

type roomUniqId map[roomID]string

type uniqIdAccount map[string]string

type rowDesc map[int]string

var mapRoomToUniqIq roomUniqId

var mapUniqIdToAccount uniqIdAccount

var mapRowDescInSheet rowDesc

func main() {
	var (
		//excelInFileName  string
		excelOutFileName string
		inputDir         string = "./In/"
	)

	mapRoomToUniqIq = make(roomUniqId)
	mapUniqIdToAccount = make(uniqIdAccount)
	mapRowDescInSheet = make(rowDesc)

	initRoomToIdzkuFromFile("Rooms.xlsx", mapRoomToUniqIq)
	initIDZhkuToElsFromFile("Accounts.xlsx ", mapUniqIdToAccount)
	inputList, _ := initInputFileList(inputDir)

	//excelInFileName = "301.xls"
	excelOutFileName = "PDTemplate.xlsx"
	for _, fileName := range inputList {
		processPlatDocFile(inputDir+fileName, excelOutFileName, mapRoomToUniqIq, mapUniqIdToAccount)
	}
}

func toUTF(inputString string) string {

	inReader := strings.NewReader(inputString)
	resReader := transform.NewReader(inReader, charmap.Windows1251.NewDecoder())
	buf, _ := ioutil.ReadAll(resReader)
	return string(buf) // строка в UTF-8
}

func processPlatDocFile(excelPD string, excelTemplate string, mapIDs roomUniqId, mapAccs uniqIdAccount) bool {
	const (
		sheetTitleRooms       = "Разделы 1-2"
		sheetTitleServices    = "Разделы 3-6"
		sheetTitlePeni        = "Неустойки"
		rowCurrentDocumentStr = "Текущий"
	)
	var (
		valStr string
		room   roomID
		//oiVolume     float64
		oiPrice                                                         float64
		oiTotalValue                                                    float64
		xlFileOutList                                                   *xlsx.File
		xlSheetOutListRooms, xlSheetOutListServices, xlSheetOutListPeni *xlsx.Sheet
	)

	fmt.Printf("Processing file %s\n", excelPD)

	xlBookPD, err := xls.Open(excelPD, "win1251")
	if err != nil {
		fmt.Println("error Open (input file)")
		return false
	}
	fmt.Println("Input file has been opened successfully")

	// проверяем есть ли уже файл с результатами, если его нет, то создаём его пустым
	if !FileExists(excelTemplate) {
		fmt.Println("Creating empty output file")
		xlFileOutList = xlsx.NewFile()

		xlSheetOutListRooms, err = xlFileOutList.AddSheet(sheetTitleRooms)
		if err != nil {
			fmt.Println("error AddSheet 1 (output file)")
			return false
		}
		fmt.Printf("Sheet %s has been added\n", sheetTitleRooms)

		xlSheetOutListServices, err = xlFileOutList.AddSheet(sheetTitleServices)
		if err != nil {
			fmt.Println("error AddSheet 2 (output file)")
			return false
		}
		fmt.Printf("Sheet %s has been added\n", sheetTitleServices)

		xlSheetOutListPeni, err = xlFileOutList.AddSheet(sheetTitlePeni)
		if err != nil {
			fmt.Println("error AddSheet 3 (output file)")
			fmt.Println(err)
			return false
		}
		fmt.Printf("Sheet %s has been added\n", sheetTitlePeni)

		xlFileOutList.Save(excelTemplate)
		fmt.Println("Output file has been created")
	}

	// теперь файл можно открывать обычным путём (он уже есть)
	xlFileOutList, err = xlsx.OpenFile(excelTemplate)

	if err != nil {
		fmt.Printf("Error on opening file %s\n", err.Error)
		return false
	}

	fmt.Println("Output file has been opened successfully")

	xlSheetPD := xlBookPD.GetSheet(0)
	if xlSheetPD == nil {
		fmt.Println("error GetSheet")
		return false
	}

	xlSheetOutListRooms = xlFileOutList.Sheet[sheetTitleRooms]
	xlSheetOutListServices = xlFileOutList.Sheet[sheetTitleServices]
	xlSheetOutListPeni = xlFileOutList.Sheet[sheetTitlePeni]

	if xlSheetOutListRooms == nil || xlSheetOutListServices == nil || xlSheetOutListPeni == nil {
		fmt.Println("Invalid structure!")
		return false
	}

	InitRowListInDocument(xlSheetPD, mapRowDescInSheet)

	// ищем период оплаты
	valStr = toUTF(xlSheetPD.Row(0).Col(0))
	periodStr, periodExists := RemovePrefixAndSuffix(valStr, "  Платежный документ (счёт) за ", " г.")
	if !periodExists {
		fmt.Printf("Period not found in '%s'\n", valStr)
		return false
	}
	fmt.Printf("period %s\n", periodStr)
	periodSlice := strings.Split(periodStr, " ")

	periodMonth := MonthNameToInt(periodSlice[0])
	periodYear, _ := strconv.Atoi(periodSlice[1])
	fmt.Printf("month = %d, year = %d\n", periodMonth, periodYear)

	// ищем номер лицевого счёта
	valStr = toUTF(xlSheetPD.Row(7).Col(6))
	accountStr, accountExists := RemovePrefixAndSuffix(valStr, "л/с ", "")
	if !accountExists {
		fmt.Printf("Account not found in '%s'\n", accountStr)
		return false
	}
	fmt.Printf("account %s\n", accountStr)

	// формируем номер платёжного документа (ГГММ+номер лицевого счёта)
	docNumber := fmt.Sprintf("%02d%02d%s", periodYear, periodMonth, accountStr)
	fmt.Printf("doc number %s\n", docNumber)

	// ищем номер квартиры
	valStr = toUTF(xlSheetPD.Row(8).Col(0))
	idx := strings.Index(valStr, "кв. ")
	room.Number, _ = strconv.Atoi(valStr[idx+6:])
	room.Type = rtLive
	fmt.Printf("room %d, ", room.Number)

	// ищем площадь
	valStr = toUTF(xlSheetPD.Row(9).Col(0))
	squareStr, squareExists := GetSubstringBetween(valStr, "Пл.:  ", " кв.м.")
	if !squareExists {
		fmt.Printf("square not found in '%s'\n", valStr)
		return false
	}
	squareVal, _ := strconv.ParseFloat(squareStr, 32)
	fmt.Printf("square %.2f, ", squareVal)

	// ищем Идентификатор помещения
	id := mapIDs[room]
	fmt.Printf("room id %s\n", id)
	// ищем Идентификатор ЖКУ
	accId := mapUniqIdToAccount[id]
	fmt.Printf("account %s\n", accId)

	// БИК и расчётный счёт
	valStr = toUTF(xlSheetPD.Row(12).Col(0))
	bankAccountStr, bankAccountExists := GetSubstringBetween(valStr, "р/счет ", " ")
	if !bankAccountExists {
		fmt.Printf("bankAccount not found in '%s'\n", valStr)
		return false
	}
	fmt.Printf("bankAccount %s\n", bankAccountStr)
	bikStr, bikExists := GetSubstringBetween(valStr, "БИК ", "")
	if !bikExists {
		fmt.Printf("BIK not found in '%s'\n", valStr)
		return false
	}
	fmt.Printf("BIK %s\n", bikStr)

	// ищем сведения о кап. ремонте
	rowVal := mapRowDescInSheet.FindRowIndex("Отчисления на капитальный ремонт")
	if rowVal < 0 {
		fmt.Printf("KapRemont info not found\n")
		return false
	}
	kapRemontRateStr := toUTF(xlSheetPD.Row(rowVal).Col(4))
	kapRemontRateVal, _ := strconv.ParseFloat(kapRemontRateStr, 32)
	fmt.Printf("KapRemont Rate %f\n", kapRemontRateVal)

	kapRemontValueStr := toUTF(xlSheetPD.Row(rowVal).Col(6))
	kapRemontValueVal, _ := strconv.ParseFloat(kapRemontValueStr, 32)
	fmt.Printf("KapRemont Value %f\n", kapRemontValueVal)

	kapRemontPereraschetStr := toUTF(xlSheetPD.Row(rowVal).Col(7))
	kapRemontPereraschetVal, _ := strconv.ParseFloat(kapRemontPereraschetStr, 32)
	if len(kapRemontPereraschetStr) > 0 {
		fmt.Printf("KapRemont Pereraschet %f\n", kapRemontPereraschetVal)
	}

	kapRemontTotalStr := toUTF(xlSheetPD.Row(rowVal).Col(8))
	kapRemontTotalVal, _ := strconv.ParseFloat(kapRemontTotalStr, 32)
	fmt.Printf("KapRemont Total %f\n", kapRemontTotalVal)

	// ищем итоговую сумму по платёжному документу
	rowItogoVal := mapRowDescInSheet.FindRowIndex("Итого")
	if rowItogoVal < 0 {
		fmt.Printf("TotalSum info not found\n")
		return false
	}
	totalDocSumStr := toUTF(xlSheetPD.Row(rowItogoVal).Col(10))
	totalDocSumVal, _ := strconv.ParseFloat(totalDocSumStr, 32)
	fmt.Printf("TotalSum %f\n", totalDocSumVal)

	// заполняем сведения о квартире в реестр платёжных документов
	//fmt.Printf("max = %d\n", xlSheetOutListRooms.MaxRow)
	/*	fmt.Printf("max2 = %d\n", xlFile2OutList.SheetCount)
		idxSheetRooms := xlFile2OutList.GetSheetIndex(sheetTitleRooms)
		fmt.Printf("id = %d\n", idxSheetRooms)*/

	// формируем строку с описанием платёжного документа
	xlRoomsRow := xlSheetOutListRooms.AddRow()
	// Идентификатор ЖКУ
	xlRoomsRow.AddCell().SetValue(accId)
	// Тип ПД
	xlRoomsRow.AddCell().SetValue(rowCurrentDocumentStr)
	// Номер платежного документа
	xlRoomsRow.AddCell().SetValue(docNumber)
	// Расчетный период (ММ.ГГГГ)
	xlRoomsRow.AddCell().SetValue(fmt.Sprintf("%02d.20%02d", periodMonth, periodYear))
	// ============= Раздел 1. Сведения о плательщике. Раздел 2. Информация для внесения платы получателю платежа (получателям платежей). =======
	// Общая площадь для ЛС
	xlRoomsRow.AddCell().SetValue("")
	// Жилая площадь
	xlRoomsRow.AddCell().SetValue("")
	// Отапливаемая площадь
	xlRoomsRow.AddCell().SetValue("")
	// Количество проживающих
	xlRoomsRow.AddCell().SetValue("")
	// Задолженность за предыдущие периоды
	xlRoomsRow.AddCell().SetValue(0)
	// Аванс на начало расчетного периода
	xlRoomsRow.AddCell().SetValue(0)
	// Учтены платежи, поступившие до указанного числа расчетного периода включительно
	xlRoomsRow.AddCell().SetValue(31)
	// БИК банка
	xlRoomsRow.AddCell().SetValue(bikStr)
	// Расчетный счет
	xlRoomsRow.AddCell().SetValue(bankAccountStr)
	// ============= Раздел 7. Расчёт размера взноса на капитальный ремонт. Раздел 8. Информация для внесения взноса на капитальный ремонт =========
	// Размер взноса на кв.м, руб.
	xlRoomsRow.AddCell().SetValue(strconv.FormatFloat(kapRemontRateVal, 'f', 2, 32))
	// Всего начислено за расчетный период, руб.
	xlRoomsRow.AddCell().SetValue(strconv.FormatFloat(kapRemontValueVal, 'f', 2, 32))
	// Перерасчеты всего, руб.
	if len(kapRemontPereraschetStr) > 0 {
		xlRoomsRow.AddCell().SetValue(strconv.FormatFloat(kapRemontPereraschetVal, 'f', 2, 32))
	} else {
		xlRoomsRow.AddCell().SetValue("")
	}
	// Льготы, субсидии, руб.
	xlRoomsRow.AddCell().SetValue("")
	// Порядок расчетов
	xlRoomsRow.AddCell().SetValue("")
	// Итого к оплате за расчетный период, руб.
	xlRoomsRow.AddCell().SetValue(strconv.FormatFloat(kapRemontTotalVal, 'f', 2, 32))
	// =========================
	// Идентификатор платежного документа
	xlRoomsRow.AddCell().SetValue("")
	// Всего
	xlRoomsRow.AddCell().SetValue(strconv.FormatFloat(totalDocSumVal, 'f', 2, 32))
	// Дополнительная информация
	xlRoomsRow.AddCell().SetValue("")

	// получаем список услуг
	rowBeginServicesVal := mapRowDescInSheet.FindRowIndex("Услуга")
	if rowBeginServicesVal < 0 {
		fmt.Println("Service list not found")
		return false
	}

	// начальные значения для Платы за содержание жилого помещения
	//oiVolume = squareVal
	oiPrice = 0.0
	oiTotalValue = 0.0

	// для каждой услуги формируем строку с её описанием
	for i := rowBeginServicesVal + 1; i < rowItogoVal; i++ {
		// Тип услуги (версия из ПД)
		serviceStr := toUTF(xlSheetPD.Row(i).Col(0))
		if strings.Compare(serviceStr, "пеня") == 0 {
			// пени надо выводить на отдельный лист
			peniStr := toUTF(xlSheetPD.Row(i).Col(10))
			peniVal, _ := strconv.ParseFloat(peniStr, 64)
			// выводим данные по пеням
			xlPeniRow := xlSheetOutListPeni.AddRow()
			// Номер платежного документа
			xlPeniRow.AddCell().SetValue(docNumber)
			// Вид начисления
			xlPeniRow.AddCell().SetValue("Пени")
			// Основания начислений
			xlPeniRow.AddCell().SetValue("Пени за просрочку коммунальный платежей")
			// Сумма, руб.
			xlPeniRow.AddCell().SetFloatWithFormat(peniVal, "0.00")
			continue
		}
		if strings.Compare(serviceStr, "текущее содержание") == 0 {
			// текущее содержание необходимо суммировать с коммунальными услугами за ОИ
			// тариф
			priceStr := toUTF(xlSheetPD.Row(i).Col(4))
			priceVal, _ := strconv.ParseFloat(priceStr, 64)
			oiPrice += priceVal
			// К оплате
			totalValueCorrStr := toUTF(xlSheetPD.Row(i).Col(10))
			totalValueCorrVal, _ := strconv.ParseFloat(totalValueCorrStr, 64)
			oiTotalValue += totalValueCorrVal
			continue
		}
		// Получаем тип услуги (версия ГИС ЖКХ)
		serviceGisName, serviceIndividual, serviceAdditional, err := ConvServiceNameToGisZhkh(serviceStr)
		if err {
			fmt.Printf("Room %d: unknown service %s\n", room.Number, serviceStr)
			return false
		}
		// Единица измерения
		//priceTypeStr := toUTF(xlSheetPD.Row(i).Col(3))
		//ConvPriceTypeToGisZjkh(priceTypeStr) // не используется!!!
		// Объём
		volumeStr := toUTF(xlSheetPD.Row(i).Col(5))
		volumeVal, _ := strconv.ParseFloat(volumeStr, 32)
		// Тариф
		priceStr := toUTF(xlSheetPD.Row(i).Col(4))
		priceVal, _ := strconv.ParseFloat(priceStr, 64)
		// Всего начислено
		totalValueStr := toUTF(xlSheetPD.Row(i).Col(6))
		totalValueVal, _ := strconv.ParseFloat(totalValueStr, 64)
		// Перерасчёт
		pereraschetStr := toUTF(xlSheetPD.Row(i).Col(7))
		pereraschetVal, _ := strconv.ParseFloat(pereraschetStr, 64)
		// К оплате
		totalValueCorrStr := toUTF(xlSheetPD.Row(i).Col(10))
		totalValueCorrVal, _ := strconv.ParseFloat(totalValueCorrStr, 64)

		if !serviceIndividual {
			// по услугам за ОИ надо всё суммировать
			oiPrice += priceVal
			oiTotalValue += totalValueVal
		}

		xlServicesRow := xlSheetOutListServices.AddRow()
		// Номер платежного документа
		xlServicesRow.AddCell().SetValue(docNumber)
		// Услуга
		xlServicesRow.AddCell().SetValue(serviceGisName)
		// индивидуальное потребление: Способ определения объемов КУ
		xlServicesRow.AddCell().SetValue("")
		// индивидуальное потребление: Объем, площадь, количество
		if serviceIndividual {
			xlServicesRow.AddCell().SetValue(strconv.FormatFloat(volumeVal, 'f', 2, 32))
		} else {
			xlServicesRow.AddCell().SetValue("")
		}
		// потребление при содержании общего имущества: Способ определения объемов КУ
		if serviceIndividual {
			xlServicesRow.AddCell().SetValue("")
		} else {
			xlServicesRow.AddCell().SetValue("Прибор учета")
		}
		// потребление при содержании общего имущества: Объем, площадь, количество
		if !serviceIndividual {
			xlServicesRow.AddCell().SetValue(strconv.FormatFloat(volumeVal, 'f', 2, 32))
		} else {
			xlServicesRow.AddCell().SetValue("")
		}
		// Тариф руб./еди-ница измерения Размер платы на кв. м, руб.
		xlServicesRow.AddCell().SetFloatWithFormat(priceVal, "0.00")
		// Всего начислено за расчетный период, руб.
		xlServicesRow.AddCell().SetFloatWithFormat(totalValueVal, "0.00")
		// Размер повышающего коэффициента
		xlServicesRow.AddCell().SetValue("")
		// Размер превышения платы, рассчитанной с применением повышающего коэффициента над размером платы, рассчитанной без учета повышающего коэффициента
		xlServicesRow.AddCell().SetValue("")
		// Перерасчеты всего, руб.
		xlServicesRow.AddCell().SetFloatWithFormat(pereraschetVal, "0.00")
		// Льготы, субсидии, руб.
		xlServicesRow.AddCell().SetValue("")
		// Порядок расчетов
		xlServicesRow.AddCell().SetValue("")
		// Норматив потребления коммунальных ресурсов: в жилых помеще-ниях
		xlServicesRow.AddCell().SetValue("")
		// Норматив потребления коммунальных ресурсов: на потребление при содержании общего имущества
		xlServicesRow.AddCell().SetValue("")
		// Текущие показания приборов учета коммунальных ресурсов: индиви-дуальных (квартир-ных)
		xlServicesRow.AddCell().SetValue("")
		// Текущие показания приборов учета коммунальных ресурсов: коллек-тивных (общедо-мовых)
		xlServicesRow.AddCell().SetValue("")
		// Суммарный объем коммунальных ресурсов в доме: в помеще-ниях дома
		xlServicesRow.AddCell().SetValue("")
		// Суммарный объем коммунальных ресурсов в доме: в целях содержания общего имущества
		xlServicesRow.AddCell().SetValue("")
		// Основания перерасчетов
		xlServicesRow.AddCell().SetValue("")
		// Сумма, руб.
		xlServicesRow.AddCell().SetValue("")
		// Сумма платы с учетом рассрочки платежа: от платы за расчетный период
		xlServicesRow.AddCell().SetValue("")
		// Сумма платы с учетом рассрочки платежа: от платы за предыдущие расчетные периоды
		xlServicesRow.AddCell().SetValue("")
		// Проценты за рассрочку: руб.
		xlServicesRow.AddCell().SetFloatWithFormat(0.0, "0.00")
		// Проценты за рассрочку: %
		xlServicesRow.AddCell().SetFloatWithFormat(0.0, "0.00")
		// Сумма к оплате с учетом рассрочки платежа и процентов за рассрочку, руб.
		xlServicesRow.AddCell().SetFloatWithFormat(totalValueCorrVal, "0.00")
		// Всего
		if serviceIndividual {
			xlServicesRow.AddCell().SetFloatWithFormat(totalValueCorrVal, "0.00")
		} else {
			xlServicesRow.AddCell().SetValue("")
		}
		// в т. ч. за ком. усл.: индивид. потребление
		if serviceIndividual && !serviceAdditional {
			xlServicesRow.AddCell().SetFloatWithFormat(totalValueCorrVal, "0.00")
		} else {
			xlServicesRow.AddCell().SetValue("")
		}
		// в т. ч. за ком. усл.: потребление при содержании общего имущества
		if !serviceIndividual && !serviceAdditional {
			xlServicesRow.AddCell().SetFloatWithFormat(totalValueCorrVal, "0.00")
		} else {
			xlServicesRow.AddCell().SetValue("")
		}
	}

	// Теперь надо вевести итоговую строку по Плате за содержание жилого помещения
	xlServicesRow := xlSheetOutListServices.AddRow()
	// Номер платежного документа
	xlServicesRow.AddCell().SetValue(docNumber)
	// Услуга
	xlServicesRow.AddCell().SetValue("Плата за содержание жилого помещения")
	// индивидуальное потребление: Способ определения объемов КУ
	xlServicesRow.AddCell().SetValue("")
	// индивидуальное потребление: Объем, площадь, количество
	xlServicesRow.AddCell().SetValue("")
	// потребление при содержании общего имущества: Способ определения объемов КУ
	xlServicesRow.AddCell().SetValue("")
	// потребление при содержании общего имущества: Объем, площадь, количество
	xlServicesRow.AddCell().SetValue("")
	// Тариф руб./еди-ница измерения Размер платы на кв. м, руб.
	xlServicesRow.AddCell().SetFloatWithFormat(oiPrice, "0.00")
	// Всего начислено за расчетный период, руб.
	xlServicesRow.AddCell().SetValue("")
	// Размер повышающего коэффициента
	xlServicesRow.AddCell().SetValue("")
	// Размер превышения платы, рассчитанной с применением повышающего коэффициента над размером платы, рассчитанной без учета повышающего коэффициента
	xlServicesRow.AddCell().SetValue("")
	// Перерасчеты всего, руб.
	xlServicesRow.AddCell().SetValue("")
	// Льготы, субсидии, руб.
	xlServicesRow.AddCell().SetValue("")
	// Порядок расчетов
	xlServicesRow.AddCell().SetValue("")
	// Норматив потребления коммунальных ресурсов: в жилых помеще-ниях
	xlServicesRow.AddCell().SetValue("")
	// Норматив потребления коммунальных ресурсов: на потребление при содержании общего имущества
	xlServicesRow.AddCell().SetValue("")
	// Текущие показания приборов учета коммунальных ресурсов: индиви-дуальных (квартир-ных)
	xlServicesRow.AddCell().SetValue("")
	// Текущие показания приборов учета коммунальных ресурсов: коллек-тивных (общедо-мовых)
	xlServicesRow.AddCell().SetValue("")
	// Суммарный объем коммунальных ресурсов в доме: в помеще-ниях дома
	xlServicesRow.AddCell().SetValue("")
	// Суммарный объем коммунальных ресурсов в доме: в целях содержания общего имущества
	xlServicesRow.AddCell().SetValue("")
	// Основания перерасчетов
	xlServicesRow.AddCell().SetValue("")
	// Сумма, руб.
	xlServicesRow.AddCell().SetValue("")
	// Сумма платы с учетом рассрочки платежа: от платы за расчетный период
	xlServicesRow.AddCell().SetValue("")
	// Сумма платы с учетом рассрочки платежа: от платы за предыдущие расчетные периоды
	xlServicesRow.AddCell().SetValue("")
	// Проценты за рассрочку: руб.
	xlServicesRow.AddCell().SetValue("")
	// Проценты за рассрочку: %
	xlServicesRow.AddCell().SetValue("")
	// Сумма к оплате с учетом рассрочки платежа и процентов за рассрочку, руб.
	xlServicesRow.AddCell().SetValue("")
	// Всего
	xlServicesRow.AddCell().SetFloatWithFormat(oiTotalValue, "# ##0,00")
	// в т. ч. за ком. усл.: индивид. потребление
	xlServicesRow.AddCell().SetValue("")
	// в т. ч. за ком. усл.: потребление при содержании общего имущества
	xlServicesRow.AddCell().SetValue("")

	// сообщение о готовности
	fmt.Printf("Room %d: processed\n", room.Number)

	// всё готово
	errSave := xlFileOutList.Save(excelTemplate)
	if err != nil {
		fmt.Printf("Error %s\n", errSave.Error())
	}

	//xlFile2OutList.SetCellStr(sheetTitleRooms, "A4", id)

	//xlFile2OutList.SaveAs("./res2.xlsx")

	return true
}

func initRoomToIdzkuFromFile(excelIDs string, mapIDs roomUniqId) {
	var isRoom bool
	var isOffice bool
	var room roomID
	var id string

	xlFile, err := xlsx.OpenFile(excelIDs)
	if err != nil {
		fmt.Println("error OpenFile")
	}
	//fmt.Printf("sheets=%d\n", len(xlFile.Sheets))
	for _, xlSheet := range xlFile.Sheets {
		if !strings.HasPrefix(xlSheet.Name, "Идентификатор") {
			continue
		}
		//fmt.Printf("sheet %s\n", xlSheet.Name)
		//fmt.Printf("rows=%d\n", len(xlSheet.Rows))
		for _, xlRow := range xlSheet.Rows {
			//fmt.Printf("row: %f %d %s\n", xlRow.Height, len(xlRow.Cells), xlRow.Cells[0].String())

			if len(xlRow.Cells) < 9 {
				continue
			}
			if !strings.HasPrefix(xlRow.Cells[0].String(), "630049") {
				continue
			}

			//fmt.Printf("%s\n", xlRow.Cells[0].String())
			isRoom = strings.Compare(xlRow.Cells[9].String(), "") != 0
			isOffice = strings.Compare(xlRow.Cells[10].String(), "") != 0
			if !isRoom && !isOffice {
				continue
			}

			id = xlRow.Cells[13].String()
			if isRoom {
				// это комната
				room.Number, _ = xlRow.Cells[9].Int()
				room.Type = rtLive
				//fmt.Printf("room %d = %s\n", room.Number, id)
			}
			if isOffice {
				// это офис
				str := xlRow.Cells[10].String()
				if strings.HasPrefix(str, "Пристройка") {
					continue
				}
				if !strings.HasPrefix(str, "оф. ") {
					continue
				}
				str = str[6:]
				pos := strings.Index(str, " (")
				if pos > 0 {
					str = str[0:pos]
				}
				room.Number, _ = strconv.Atoi(str)
				room.Type = rtOffice
				//fmt.Printf("office %d = %s\n", room.Number, id)
			}
			mapIDs[room] = id
		}
	}
	fmt.Printf("Reading %d rooms from file\n", len(mapIDs))
}

func initIDZhkuToElsFromFile(excelIDs string, mapIDs uniqIdAccount) {
	var acc string
	var id string

	xlFile, err := xlsx.OpenFile(excelIDs)
	if err != nil {
		fmt.Println("error OpenFile")
	}
	//fmt.Printf("sheets=%d\n", len(xlFile.Sheets))
	for _, xlSheet := range xlFile.Sheets {
		if strings.Compare(xlSheet.Name, "Шаблон экспорта ЕЛС") != 0 {
			continue
		}
		//fmt.Printf("sheet %s\n", xlSheet.Name)
		//fmt.Printf("rows=%d\n", len(xlSheet.Rows))
		for _, xlRow := range xlSheet.Rows {
			//fmt.Printf("row: %f %d %s\n", xlRow.Height, len(xlRow.Cells), xlRow.Cells[0].String())
			if strings.Compare(xlSheet.Name, "Номер ЛС") == 0 {
				continue
			}
			if len(xlRow.Cells) < 5 {
				continue
			}

			//fmt.Printf("%s\n", xlRow.Cells[0].String())

			acc = xlRow.Cells[2].String()
			id = xlRow.Cells[3].String()
			mapIDs[id] = acc
		}
	}
	fmt.Printf("Reading %d accounts from file\n", len(mapIDs))
}

func RemovePrefixAndSuffix(s, prefix, suffix string) (resStr string, found bool) {
	found = false
	resStr = ""

	if !strings.HasPrefix(s, prefix) || !strings.HasSuffix(s, suffix) {
		return
	}
	found = true
	resStr = strings.TrimPrefix(s, prefix)
	resStr = strings.TrimSuffix(resStr, suffix)
	return
}
func GetSubstringBetween(s, before, after string) (resStr string, found bool) {
	found = false
	resStr = ""

	p := strings.Index(s, before)
	if p < 0 {
		return
	}
	s = s[p+len(before):]
	if len(after) > 0 {
		p = strings.Index(s, after)
		if p < 0 {
			return
		}
		resStr = s[0:p]
		found = true
	} else {
		resStr = s
		found = true
	}
	return
}

// Преобразует название месяца (на русском) в номер месяца
func MonthNameToInt(month string) int {
	var months = []string{"Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август"}

	for p, v := range months {
		if v == month {
			return p + 1
		}
	}
	return -1
}

func InitRowListInDocument(sheet *xls.WorkSheet, mapList rowDesc) bool {
	for i := 0; i < int(sheet.MaxRow); i++ {
		mapList[i] = toUTF(sheet.Row(i).Col(0))
	}
	return true
}

func (mapList rowDesc) FindRowIndex(rowTitle string) int {
	for p, v := range mapList {
		if strings.Compare(v, rowTitle) == 0 {
			return p
		}
	}
	return -1
}

/*func ConvPriceTypeToGisZjkh(priceType string) (resType string, err bool) {
	var (
		origType = []string{"м2", "Квт.ч", "м3"}
		gisType  = []string{""}
	)
	return
}*/

func ConvServiceNameToGisZhkh(serviceName string) (resName string, individual bool, additional bool, err bool) {
	var (
		origServiceName = []string{"охрана", "домофон", "видеодомофон",
			"холодное водоснабжение", "горячее водоснабжение", "водоотведение", "электроэнергия",
			"электроэнергия на содерж. ОИ", "горячая вода на содерж.  ОИ", "холодная вода на содерж. ОИ"}
		gisServiceName = []string{"Оплата охранных услуг", "Запирающее устройство (ЗУ)", "Видеонаблюдение",
			"Холодное водоснабжение", "Горячее водоснабжение", "Водоотведение", "Электроснабжение",
			"Электрическая энергия", "Горячая вода", "Холодная вода"}
		isIndividual = []bool{true, true, true,
			true, true, true, true,
			false, false, false}
		isAdditional = []bool{true, true, true,
			false, false, false, false,
			false, false, false}
		serviceSimpleName string
	)

	resName = ""
	individual = true
	additional = false
	err = true

	pos := strings.Index(serviceName, " (")
	if pos > 1 {

		serviceSimpleName = serviceName[0:pos]
	} else {
		serviceSimpleName = serviceName
	}

	for p, v := range origServiceName {
		if v == serviceSimpleName {
			resName = gisServiceName[p]
			individual = isIndividual[p]
			additional = isAdditional[p]
			err = false
			break
		}
	}
	return
}

func initInputFileList(inputDir string) (fileList map[int]string, err bool) {

	count := 0
	fileList = make(map[int]string, 100)
	err = false

	// читаем список файлов из входного каталога
	dirEntries, _ := ioutil.ReadDir(inputDir)
	for _, val := range dirEntries {
		if val.IsDir() {
			continue
		}
		if strings.Compare(filepath.Ext(val.Name()), ".xls") != 0 {
			continue
		}
		fileList[count] = val.Name()
		count++
	}
	fmt.Printf("Finding %d input files\n", len(fileList))
	return
}

func FileExists(fileName string) bool {
	if _, err := os.Stat(fileName); err != nil {
		if os.IsNotExist(err) {
			return false
		}
	}
	return true
}
