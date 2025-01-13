// import './App.css';
import './styles.css';
import { NavLink, Link, Route, Routes } from 'react-router-dom';
import Home from './pages/Home';

import NotFound from './pages/NotFound';
// import { BookLayout } from './BookLayout';
import { BookRoutes } from './BookRoutes';

function App() {
  return (
    <>
      <nav>
        <ul>
          <li>
            <NavLink to="/">Home</NavLink>
          </li>
          <li>
            <NavLink end to="/books">
              Books
            </NavLink>
          </li>
        </ul>
      </nav>

      <Routes>
        <Route path="/" element={<Home />} />
        <Route path="/books/*" element={<BookRoutes />} />
        <Route path="*" element={<NotFound />} />
      </Routes>
    </>
  );
}

export default App;

{
  /* <Route path="/books" element={<BookLayout />}>
          <Route index element={<BookList />} />
          <Route path=":id" element={<Book />} />
          <Route path="new" element={<NewBook />} />
        </Route> */
}

{
  /* <Routes location="/books">
        <Route path="/books" element={<h1>Extra Content</h1>} />
      </Routes> */
}

// {({ isActive }) => {
//   return isActive ? 'Active Home' : 'Home';
// }}

// style={({ isActive }) => {
//   return isActive ? { color: 'red' } : {};
// }}
