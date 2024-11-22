package main

import (
	"fmt"
	"log"
	"sort"
	"strings"
	"time"

	"github.com/gocolly/colly"
	"github.com/xuri/excelize/v2"
)

type Product struct {
	Name            string
	Characteristics map[string]string
}

func main() {

	listCollector := colly.NewCollector(
		colly.UserAgent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"),
	)
	listCollector.SetRequestTimeout(2 * time.Second)
	i := 1
	var products []Product
	listCollector.OnHTML(".product-item", func(h *colly.HTMLElement) {
		//Сохраняем ссылку на каждый товар вне зависимости от вложенности
		link := h.ChildAttr(".product__listing.product__grid .details .name", "href")
		//Сохраняем имя товара
		name := h.ChildAttr(".product__listing.product__grid .details .name", "title")
		//Сохраняем значение атрибута, которое есть только у товаров, которые не имеют вложенности (ссылка на 1 товар, а не на несколько)
		//Если товар 1, то значение не будет пустым, если их несколько, то значение будет пустой строкой
		//В зависимости от вложенности товары внутри будут обрабатываться разными функциями
		enclosureType := h.ChildText(".product__listing.product__grid .details .order-number")

		fmt.Println("Парсинг ", i)
		i++
		if enclosureType == "" { //Товаров несколько, используем функция для поиска всех карточек во вложении
			product := childList(link, name)
			products = append(products, product...)
		} else { //Товар один, используем функция для поиска характеристик
			product := Product{
				Name:            name,
				Characteristics: make(map[string]string),
			}
			characteristic, err := findDetails(link)
			if err != nil {
				log.Println("Ошибка при сборе характеристик для: ", link, ":", err)
			} else {
				product.Characteristics = characteristic
				products = append(products, product)

			}
		}
	})

	if err := listCollector.Visit("https://shop.hettich.com/ru_RU/%D0%9F%D0%B5%D1%82%D0%BB%D0%B8/c/group2264954678785?q=%3Aorder-asc&page=0"); err != nil {
		panic(err)

	}

	//Создание эксель файла
	f := excelize.NewFile()
	sheetName := "Products"
	f.SetSheetName("Sheet1", sheetName)

	//Собираем все уникальные ключи характеристик
	keys := getAllKeys(products)

	//Записываем заголовки (название товаров) в первую строку
	f.SetCellValue(sheetName, "A1", "Название")
	f.SetCellValue(sheetName, "B1", "Характеристики")
	for col, product := range products {
		cell := fmt.Sprintf("A%d", 3+col)
		f.SetCellValue(sheetName, cell, product.Name)
	}

	// Заполнение эксель файла характеристиками
	row := 0
	for ind, key := range keys {
		f.SetCellValue(sheetName, fmt.Sprintf("%c2", 'B'+ind), key)

		for col, product := range products {
			value := product.Characteristics[key]
			cell := fmt.Sprintf("%c%d", 'B'+row, 3+col)
			f.SetCellValue(sheetName, cell, value)
		}

		row++

	}
	//Сохранение эксель файла
	fileName := "products.xlsx"
	if err := f.SaveAs(fileName); err != nil {
		log.Fatalf("Ошибка сохранения Excel файла: %s", err)
	}
	fmt.Printf("Данные успешно записаны в файл %s\n", fileName)

}

// Функция для нахождения всех характеристик каждого товара
func findDetails(link string) (map[string]string, error) {
	characteristic := make(map[string]string)
	productCollector := colly.NewCollector(
		colly.UserAgent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"),
	)
	productCollector.SetRequestTimeout(2 * time.Second)

	productCollector.OnHTML(".table", func(r *colly.HTMLElement) {

		details := []string{}
		r.ForEach(".product-classifications table.table td", func(_ int, t *colly.HTMLElement) {
			text := strings.TrimSpace(t.Text)
			details = append(details, text)

		})
		
		for i := 0; i <= len(details)-2; {
			characteristic[details[i]] = details[i+1]
			i = i + 2

		}

	})
	productCollector.OnError(func(r *colly.Response, err error) {
		log.Println("Ошибка при парсинге ", link, ":", err)
	})

	err := productCollector.Visit("https://shop.hettich.com" + link)
	if err != nil {
		return nil, err
	}
	return characteristic, nil
}

// Функция для нахождения карточек товара во вложении
func childList(url, name string) []Product {
	var products []Product
	childListCollector := colly.NewCollector(
		colly.UserAgent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"),
	)
	childListCollector.SetRequestTimeout(2 * time.Second)

	childListCollector.OnHTML("table.dataTable tbody tr", func(h *colly.HTMLElement) {
		link := h.ChildAttr(".datatable-slim-style a", "href")

		product := Product{
			Name:            name,
			Characteristics: make(map[string]string),
		}

		products = append(products, product)

		characteristics, err := findDetails(link)
		if err != nil {
			log.Println("Ошибка при сборе характеристик для", link, ":", err)
		} else {
			products[len(products)-1].Characteristics = characteristics
		}

	})
	err := childListCollector.Visit("https://shop.hettich.com" + url)
	if err != nil {
		log.Println("Ошибка при переходе по ссылке: ", err)
		return nil
	}

	return products
}

// Функция для нахождения всех индивидуальных характеристик товара
func getAllKeys(products []Product) []string {
	keySet := make(map[string]struct{})
	for _, product := range products {
		for key := range product.Characteristics {
			keySet[key] = struct{}{}
		}
	}
	keys := make([]string, 0, len(keySet))
	for key := range keySet {
		keys = append(keys, key)
	}
	sort.Strings(keys)
	return keys
}
