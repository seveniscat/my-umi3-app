import { handleFile, test } from './tool';

export default function IndexPage() {
  return (
    <div>
      <h1>Hello World</h1>

      <input type='file' accept='.xls,.xlsx' onChange={(k) => {
        handleFile(k.target.files?.[0])
      }}></input>

      <button onClick={test}>Test</button>
    </div>
  );
}
