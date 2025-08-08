let oldArray = [4, 5, 6];

let myArray = [1,2,3,...oldArray.map(n => {
    return (n, n*2)
})];

console.log(myArray);