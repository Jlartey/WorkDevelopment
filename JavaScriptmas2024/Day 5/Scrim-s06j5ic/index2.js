const snowGlobe = document.querySelector('.snow-globe');
let snowflakeCount = 0;
let isSnowing = false;

function createSnowflake() {
  if (!isSnowing) return; // Only create snowflakes when snowing is active

  // Create snowflake element
  const snowflake = document.createElement('div');
  snowflake.classList.add('snowflake');

  // Randomize snowflake type (every 25th is a snowman)
  snowflakeCount++;
  snowflake.innerHTML = snowflakeCount % 25 === 0 ? '☃️' : '❄️';

  // Randomize properties
  const startX = Math.random() * 100; // Percentage across snow globe width
  const size = Math.random() * 20 + 10; // 10px to 30px
  const duration = Math.random() * 5 + 3; // 3s to 8s
  const sway = Math.random() * 20 - 10; // -10px to 10px horizontal sway

  // Apply styles
  snowflake.style.left = `${startX}%`;
  snowflake.style.fontSize = `${size}px`;
  snowflake.style.animationDuration = `${duration}s`;
  snowflake.style.setProperty('--sway', `${sway}px`); // Custom property for animation

  // Append to snow globe
  snowGlobe.appendChild(snowflake);

  // Remove snowflake after animation ends to avoid clutter
  setTimeout(() => {
    snowflake.remove();
  }, duration * 1000);
}

// Toggle snowing with a button
const startButton = document.createElement('button');
startButton.textContent = 'Shake the Globe!';
document.body.appendChild(startButton);

startButton.addEventListener('click', () => {
  if (!isSnowing) {
    isSnowing = true;
    snowGlobe.classList.add('shake');
    const snowInterval = setInterval(createSnowflake, 200); // Faster snow at first

    // Slow down and stop snow after 10 seconds
    setTimeout(() => {
      clearInterval(snowInterval);
      setInterval(createSnowflake, 1000); // Less frequent snow
      setTimeout(() => {
        isSnowing = false;
        snowGlobe.classList.remove('shake');
      }, 5000); // Stop after 5 more seconds
    }, 10000);
  }
});
