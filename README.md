Learn Apache POI XSSF
=====================

This project includes example code for uploading and downloading Excel Spreadsheets
using Apache POI XSSF APIs and data from a database.

There are examples of spreadsheet parsing using both SAX and DOM algorithms, and
saving data using both JDBC and JPA.

To build the application, run:

```
./mvnw clean package
```

To run the application:

```
java -jar target/learn-apache-poi-xssf-0.0.1-SNAPSHOT.jar
```

Using your browser, use this link:

```
http://localhost:8080/index.html
```

There are four sections to the application:

1) Download random test data

2) Upload a spreadhsheet with controls for various aspects

3) Download data previously uploaded

4) Truncate database

When using the application, times to perform operations are logged by the application.

## Insert Performance

Here are example times to upload a spreadsheet with one million rows:

### SAX / JDBC / Inserts / Batch Size: 1000 (BEST)
```
Time to save 1000000 widgets: StopWatch 'sax-jdbc': running time (millis) = 16802; [] took 16802 = 100%
```

### SAX / JPA / Inserts / Batch Size: 1000
```
Time to save 1000000 widgets: StopWatch 'sax-jpa': running time (millis) = 25359; [] took 25359 = 100%
```

### DOM / JDBC / Inserts / Batch Size: 1000
```
Time to save 1000000 widgets: StopWatch 'dom-jdbc': running time (millis) = 30494; [] took 30494 = 100%
```

### DOM / JPA / Inserts / Batch Size: 1000 (WORST)
```
Time to save 1000000 widgets: StopWatch 'dom-jpa': running time (millis) = 33701; [] took 33701 = 100%
```

## Update Performance

Here are example times to upload a spreadsheet with one million rows:

### SAX / JDBC / Updates / Batch Size: 1000 (BEST)
```
Time to save 1000000 widgets: StopWatch 'sax-jdbc': running time (millis) = 23841; [] took 23841 = 100%
```

### SAX / JPA / Updates / Batch Size: 1000
```
Time to save 1000000 widgets: StopWatch 'sax-jpa': running time (millis) = 38525; [] took 38525 = 100%
```

### DOM / JDBC / Updates / Batch Size: 1000
```
Out of Memory Error
```

### DOM / JPA / Updates / Batch Size: 1000 (WORST)
```
Out of Memory Error
```

### TODO

* Add more fields to Widget (e.g. expires, radioactive flag, weight, age, color)
* Add more validation logic (work with cell type?)
* Review schema / indexing
* Verify transaction configuration
* Consider UI enhancements
* Consider a profile that does not use an in memory database

### Conclusion

I think you should just use SAX and JDBC. But can I get JPA to perform almost as good?

  