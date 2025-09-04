const snowGlobe = document.querySelector('.snow-globe')

function createSnowflake() {
/* 
Challenge:
1. Write JavaScript to create a snowflake and make it fall inside the snow globe. The snowflake should have a random starting position, animation duration, and size.
2. See index.css
*/ 
// Create snowflake element
    const snowflake = document.createElement('div');
    snowflake.classList.add('snowflake');
    snowflake.innerHTML = '❄️';

    // Randomize properties
    const startX = Math.random() * 100; // Percentage across snow globe width
    const size = Math.random() * 20 + 10; // 10px to 30px
    const duration = Math.random() * 5 + 3; // 3s to 8s

    // Apply styles
    snowflake.style.left = `${startX}%`;
    snowflake.style.fontSize = `${size}px`;
    snowflake.style.animationDuration = `${duration}s`;

    // Append to snow globe
    snowGlobe.appendChild(snowflake);
}

setInterval(createSnowflake, 3000) // Let's create a snowflake every 100 milliseconds!

/* Stretch goals: 
- Give some variety to your snowflakes, so they are not all the same. Perhaps every 25th one could be a snowman ☃️?
- Remove each snowflake after a set time - this will stop the scene from being lost in a blizzard!
- Add a button that makes the snow start falling, it could trigger a CSS-animated shake of the snow globe. Then make the snow become less frequent until it slowly stops - until the button is pressed again.  
- Change the direction of the snowflakes so they don’t all fall vertically.
- Make the style your own! 
*/