# VBAScripting

## Summary of the project

To refactor the existing VBA script to run faster.

## Situation
Steve uses VBA script to analyze the stock data. However current script is not effiecient and takes longer time when used in larger data.
Use content learned thorughout this module to refactor the code faster.  Orignally code took 1.49 and 0.05 seconds to run in 2017 and 2018 respectfully. 


<img width="272" alt="Orignalruntime2017" src="https://user-images.githubusercontent.com/93223274/140663018-744e0d01-5c2e-47e0-8150-932c88207113.png">

<img width="263" alt="Orignalruntime2018" src="https://user-images.githubusercontent.com/93223274/140663019-154c2e53-8cb8-48f7-b512-1c8bee98f028.png">


### Refactoring

Orignal code lops through 12 stockes. It picks one and loops through the entire dataset. Every time macro is run it loops thorugh 3000 lines 12 times. This can be minimized by the use of stock ticker. Refactor code creates three new arrays that holds tickervolumnes, tickerstartingprice and tickerendingprice. Below is the results of the refactored code. 


<img width="267" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/93223274/140663025-052db7ba-1b15-4d15-9094-b3c9337e9e7d.png">

<img width="257" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/93223274/140663026-6b34baa3-f674-44d7-99c8-00aea2e7a134.png">


### The Data
The rate of the return is much better in the 2017 than in 2018. There are some stocks like TERP, RUN and ENPH that are not much affected throughout the time. Through the invesment prepestive some of the stocks like RUN, ENPH that has higher total daily volumne and rate of return. These are relatively safer stocks to invest in. 


>![2017Summary](https://user-images.githubusercontent.com/93223274/140663023-a9803842-89da-4ce8-859e-e63372bc0308.png)>
![2018Summary](https://user-images.githubusercontent.com/93223274/140663024-1239afe5-c273-42a9-8a7a-1f903e9fadb4.png)>

### Refactoring as a practise
Refactoring is the key procedural part of the product development. Often times all the programming languages keeps evolving and they include packages and subpackes that makes tasking efficient with fewer lines of code. Refactoring is not just good for efficiency it also improves the readability which can be useful in the tranfer of the knowledge. Often time when working in the big team this can save time. There are times when one should think hard before refactoring as it can sometime effect in the testing outcome. One should know the impact of the refactoring to the exitisting pipeline before actually doing it. 
