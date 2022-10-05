# VAR-Model-Prediction
Automated a Vector Auto Regression Model in Python in order to predict price movements of commodities based on effect from other countries with a time lag.

The file 'S1-Proizvodi.xlsx' is the input database containing 14 tabs 1 for each commodity and the monthly price changes over a 10 year time frame for the available countries.

The 'Rezultati_Proizvodi.xlsx' is the output generated from the Python script, where each tab contains the following info: 
  * The p-value of every country's ad-fuller test used to check stationarity
  * The optimal time lag between 1 and 6 months used for the VAR model based on the AIC,BIC,FPE HQIC scores.
  * The predictions for Macedonia prices in column AG as Macedonia_predicted.




