package main

import (

	"fmt"
	"log"
	"net/http"
	"html/template"
	"os"
	"io"
	"strings"
	"net/smtp"
	"strconv"
	"encoding/json"
	_"time"
	_"path/filepath"
	"os/exec"
	"runtime"
	_"github.com/Luzifer/go-openssl"
	_"github.com/gabriel-vasile/mimetype"
	"github.com/360EntSecGroup-Skylar/excelize"
)

type column struct {
	ColName, ColId string
}

const MAX_UPLOAD_SIZE = 2048 * 1024 // 2MB

func main(){

	http.HandleFunc("/index", index)
	http.HandleFunc("/help", help)
	http.HandleFunc("/uploadHandler", uploadHandler)
	http.HandleFunc("/mailSender", mailSender)


	http.Handle("/static/", http.StripPrefix("/static/", http.FileServer(http.Dir("static"))))
	fmt.Println("Pour accéder à la plate-forme, veuillez vous connecter au site: localhost:8080/index" )	
	http.ListenAndServe(":8080", nil)
}

func index(w http.ResponseWriter, r *http.Request) {

	render(w,"templates/nav.html", "renderingnavBar")
	tpl := template.Must(template.ParseFiles("templates/index.html"))
	err := tpl.ExecuteTemplate(w, "index.html",nil)
	if err != nil {
		log.Println(err)
		http.Error(w, "Internal server error !", http.StatusInternalServerError)
		return
	}
}

func help(w http.ResponseWriter, r *http.Request) {
	render(w, "templates/nav.html", "renderingnavBar")
	tpl := template.Must(template.ParseFiles("templates/help.html"))
	err := tpl.ExecuteTemplate(w, "help.html", nil)
	if err != nil {
		log.Println(err)
		http.Error(w, "Internal server error !", http.StatusInternalServerError)
		return
	}
}

func mailSender(w http.ResponseWriter, r *http.Request) {

	if r.Method != "POST" {
		http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
		return
	}

	Object :=  r.PostFormValue("Object")
	Content := "<html><body>"+r.PostFormValue("Content")+"</body></html>"
	NoReply := r.PostFormValue("NoReply")
	ColumnsMapJSON := r.PostFormValue("ColumnsMapJSON")
	FileName := r.PostFormValue("FileName")
	SheetName := r.PostFormValue("SheetName")
	MailColumn :=  r.PostFormValue("MailColumn")
	from := r.PostFormValue("email")
	pass := r.PostFormValue("password")

	fmt.Println("from, pass:", from, pass)

	var ColumnsMapTab [] column

	b := []byte(ColumnsMapJSON)
	err := json.Unmarshal(b, &ColumnsMapTab)
	if err != nil {
		http.Error(w, "not valid JSON", http.StatusMethodNotAllowed)
	}
	
	var results struct {
		NbMailSent string 
	}

	f, err := excelize.OpenFile("./uploads/"+FileName)
    if err != nil {
        fmt.Println(err)
        return
    }


    mime := "MIME-version: 1.0;\nContent-Type: text/html; charset=\"UTF-8\";\n\n"
    auth := smtp.PlainAuth("", from, pass, "smtp.gmail.com")
    nbMailSent := 0

    var val, dest string
    rows, _ := f.GetRows(SheetName)

    newContent := ""

    for ligne := 2 ; ligne <= len(rows) ; ligne++ {
    	
    	newContent = Content
    	dest, _ = f.GetCellValue(SheetName, MailColumn+strconv.Itoa(ligne))
    	for i := 0; i < len(ColumnsMapTab); i++ {
    		if strings.Contains(Content, "{{"+ColumnsMapTab[i].ColName+"}}") {
    			val, _ = f.GetCellValue(SheetName, ColumnsMapTab[i].ColId+strconv.Itoa(ligne))
    			newContent = strings.ReplaceAll(newContent, "{{"+ColumnsMapTab[i].ColName+"}}", val)
    		}
    	}

   	    if NoReply == "on" {
			newContent = newContent + "<br><p>Ce mail est diffusé automatiquement, veuillez ne pas répondre </p>"	
		}

		msg := []byte("Subject: " + Object + "\n" + mime + "<html><body>"+ newContent +"</body></html>")

		err := smtp.SendMail("smtp.gmail.com:587", auth, from, []string{dest}, msg)
		if err != nil {
			fmt.Println("smtp error: %s", err)
			return	
		} else {
	    	nbMailSent++
	    }
	}
 
    results.NbMailSent = strconv.Itoa(nbMailSent)

	render(w, "templates/nav.html", "renderingnavBar")
	tpl := template.Must(template.ParseFiles("templates/successfull.html"))
	err = tpl.ExecuteTemplate(w, "successfull.html", results)
	if err != nil {
		log.Println(err)
		http.Error(w, "Internal server error !", http.StatusInternalServerError)
		return
	}
}

func send(subject, body, dest, from, pass string) {
	mime := "MIME-version: 1.0;\nContent-Type: text/html; charset=\"UTF-8\";\n\n"
	msg := []byte("Subject: " + subject + "\n" + mime + "<html><body>"+ body +"</body></html>")

	err := smtp.SendMail("smtp.gmail.com:587", smtp.PlainAuth("", from, pass, "smtp.gmail.com"), from, []string{dest}, msg)
	if err != nil {
		fmt.Println("smtp error: %s", err)
		return
	}
}

func uploadHandler(w http.ResponseWriter, r *http.Request) {

    var results struct {
		Columns	[] column
		MailCol string
		ColumnsMapJSON string
		FileName string
		SheetName string
		SheetList map[int]string
		EmptySheet string
	}

	getfilename := 	""
	getSheetName := ""
	SheetIndex := 0

	if r.Method == "POST" {
		
		r.Body = http.MaxBytesReader(w, r.Body, MAX_UPLOAD_SIZE)
		if err := r.ParseMultipartForm(MAX_UPLOAD_SIZE); err != nil {
			http.Error(w, "The uploaded file is too big. Please choose an file that's less than 1MB in size", http.StatusBadRequest)
			return
		}

		// The argument to FormFile must match the name attribute of the file input on the frontend
		file, fileHeader, err := r.FormFile("myfile")
		if err != nil {
			http.Error(w, err.Error(), http.StatusBadRequest)
			return
		}

		defer file.Close()
		
		/*filetype := http.DetectContentType(buff)
		if filetype != "application/octet-stream" {
			http.Error(w, "The provided file format is not allowed. Please upload a xslx or xls file", http.StatusBadRequest)
			return
		}*/

		//fmt.Println(r.FormValue("email"))

		// Create the uploads folder if it doesn't  already exist
		err = os.MkdirAll("./uploads", os.ModePerm)
		if err != nil {
			http.Error(w, err.Error(), http.StatusInternalServerError)
			return
		}

		// Create a new file in the uploads directory
		dst, err := os.Create(fmt.Sprintf("./uploads/%s", fileHeader.Filename))
		if err != nil {
			http.Error(w, err.Error(), http.StatusInternalServerError)
			return
		}

		defer dst.Close()

		// Copy the uploaded file to the filesystem at the specified destination
		_ , err = io.Copy(dst, file)
		if err != nil {
			http.Error(w, err.Error(), http.StatusInternalServerError)
			return
		}
		getfilename = fileHeader.Filename
	
	} else {
		if r.Method == "GET" {
			query := r.URL.Query()
			getfilename = query.Get("FileName")
			getSheetName = query.Get("SheetName")

		}
	}

	results.FileName = getfilename

	f, err := excelize.OpenFile("./uploads/"+getfilename)
    if err != nil {
        fmt.Println(err)
        return
    }

    results.SheetList = f.GetSheetMap()
    if r.Method == "GET" {
    	SheetIndex = f.GetSheetIndex(getSheetName)
    }

	var elt column

    rows, _ := f.GetRows(f.GetSheetName(SheetIndex))
    results.SheetName = f.GetSheetName(SheetIndex)

    
    if len(rows) > 0 { 
	    for i, row := range rows[0] {
	    	if row != "" {
	    		elt.ColName = row
	    		elt.ColId = colItoA(i)
	    		//fmt.Println("index: ",elt.ColId)
	    		if strings.Contains(row, "mail") {
	    			results.MailCol = row
	    		}
	        	results.Columns = append(results.Columns, elt)
	        }
	    }
	    results.EmptySheet = "N"
	} else {
		results.EmptySheet = "Y"	
	}

    ColJSON, err := json.Marshal(&results.Columns)
    results.ColumnsMapJSON = string(ColJSON)

	render(w,"templates/nav.html", "renderingnavBar")
	tpl := template.Must(template.ParseFiles("templates/mailEditor.html"))
	err = tpl.ExecuteTemplate(w, "mailEditor.html", results)
	if err != nil {
		log.Println(err)
		http.Error(w, "Internal server error !", http.StatusInternalServerError)
		return
	}
}

func colItoA(index int) string {
	Alphabet := "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	reminder := index % 26
	result := index / 26
	str := string(Alphabet[reminder : reminder+1])
	for result > 0 {
		reminder = result % 26
		result = result / 26
		str = string(Alphabet[reminder-1 : reminder])+str
	}
	return str
}

func render(w http.ResponseWriter, filename string, data interface{}) {
  tmpl, err := template.ParseFiles(filename)
  if err != nil {
    http.Error(w, err.Error(), http.StatusInternalServerError)
  }
  if err := tmpl.Execute(w, data); err != nil {
    http.Error(w, err.Error(), http.StatusInternalServerError)
  }
}


/*func decrypt(crypt string) string {
    encrypted := "ENCRYPTED_STRING_HERE"
    secret := "394812730425442A382D2F423F452848"

    o := openssl.New()

    dec, err := o.DecryptBytes(secret, []byte(encrypted), openssl.DigestMD5)
    if err != nil {
        fmt.Printf("An error occurred: %s\n", err)
    }

    fmt.Printf("Decrypted text: %s\n", string(dec))

    return string(dec)
} */

func open(url string) {
    var err error

	switch runtime.GOOS {
	case "linux":
		err = exec.Command("xdg-open", url).Start()
	case "windows":
		err = exec.Command("rundll32", "url.dll,FileProtocolHandler", url).Start()
	case "darwin":
		err = exec.Command("open", url).Start()
	default:
		err = fmt.Errorf("unsupported platform")
	}
	if err != nil {
		log.Fatal(err)
	}
}