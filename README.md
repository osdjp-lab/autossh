# AutoSSH

## Softship program startup requirements

- Open "Update booking" tab
- Booking number set as first column
- "Update booking" in searched state overwise initial state is default query mode

## Possible improvements

- State tracking via pixel / area reads in place of hard coded wait times
- Extend to handle haulier assignment

## Implementations route

1. Finds window checks state (switches to query mode from result mode if necessary) and enters job number, opens job and enters costs tab
2. Takes screenshot of costs
3. Calculates diff

