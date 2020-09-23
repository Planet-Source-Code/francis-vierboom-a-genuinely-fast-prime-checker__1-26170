<div align="center">

## A Genuinely Fast Prime Checker


</div>

### Description

This program is faster than any prime checker I've seen here. 2 main reasons

1)this only checks for factors up to the square root of the number, not half the number like most programs.

2)it only checks for prime factors. (number theory 101 - all factors are primes or products of primes. no point in checking to see if 12's a factor if you've already checked 2 or 3.)

<p>

obviously this means 'training' the program so it knows lots of primes. recommended: before you get started, use the 'find primes' up to 100000 (40 sec on a p200). this will dramatically speed up all prime checks up to 10^10. if you train it up to 1000000 (3min 24sec) then it will be much faster all the way up to 10^12.

<p>

the program holds all the prime numbers that have been found so far in both an array and a file, so after a while it begins to eat memory. i haven't reallly tested that; its a warning just in case though.

<p>

the best i've seen is a checker that decided 1000000000000091 was a prime in 5 min and 5 sec (Mike Frey at http://www.planet-source-code.com/xq/ASP/txtCodeId.7168/lngWId.1/qx/vb/scripts/ShowCode.htm). This one does it in 21.1 seconds (having trained it up to 4.6 million). I'm sure theres still some tweaking to do though...
 
### More Info
 
watch out, as you'd expect it takes up a fair bit of processing power. add some doevents if you're worried.


<span>             |<span>
---                |---
**Submitted On**   |2000-11-21 02:45:56
**By**             |[Francis Vierboom](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/francis-vierboom.md)
**Level**          |Intermediate
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[A Genuinel247308152001\.zip](https://github.com/Planet-Source-Code/francis-vierboom-a-genuinely-fast-prime-checker__1-26170/archive/master.zip)

### API Declarations

```
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
```





