## Serial Logger

A tool to receive serial data from a microcontroller and save it to a textfile or Microsoft Excel. Log contiously or request data by sending a serial message to the microcontroller which is programmed to respond (see code below).


### Excel usage:
- open a worksheet and select the cell you want to store the data in
- select a option in which direction the cell will move after a complete datapacket has been received
- Choose a delimiter to show where the Dataset ends (e.g. in '33;34;35#' the '#' shows the end of the dataset)
- Choose a datapoint delimiter to differentiate the datapoints (e.g. in '33;34;35#' the ';' shows the end of the datapoint)

### Arduino example code:
```
void setup() {
  // start serial port at 9600 bps:
  Serial.begin(9600);
  while (!Serial) {
    ; // wait for serial port to connect. Needed for native USB port only
  }
  establishContact();  // send a byte to establish contact until receiver responds
}

void loop() {
  // if we get a valid byte, send data:
  if (Serial.available() > 0) {
    Serial.println("32;33;34;");
  }
}

void establishContact() {
  while (Serial.available() <= 0) {
    Serial.println("1;2;3;");
    delay(1000);
  }
}
```

