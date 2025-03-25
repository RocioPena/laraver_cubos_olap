<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Consulta CLUES</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body class="bg-light py-5">
<div class="container">
    <h2 class="text-center mb-4">Consulta dinámica de CLUES</h2>

    @if(session('error'))
        <div class="alert alert-danger">{{ session('error') }}</div>
    @endif

    <form method="POST" action="{{ route('consultar.clues') }}" class="card p-4 shadow mb-4">
        @csrf
        <div class="mb-3">
            <label for="clues" class="form-label">CLUES</label>
            <input type="text" name="clues" id="clues" class="form-control" required value="{{ old('clues', $clues ?? '') }}">
        </div>

        <div class="mb-3">
            <label for="variables" class="form-label">Variables (claves separadas por coma)</label>
            <input type="text" name="variables" id="variables" class="form-control" required value="{{ old('variables', $variables ?? '') }}">
        </div>

        <div class="mb-3">
            <label for="modo" class="form-label">Modo de encabezado</label>
            <select name="modo" id="modo" class="form-select">
                <option value="nombre" {{ ($modo ?? '') === 'nombre' ? 'selected' : '' }}>Solo nombre</option>
                <option value="codigo+nombre" {{ ($modo ?? '') === 'codigo+nombre' ? 'selected' : '' }}>Código + nombre</option>
            </select>
        </div>

        <button type="submit" class="btn btn-primary w-100">🔍 Consultar</button>
    </form>

    @isset($datos)
        <h4 class="mb-3">Resultados</h4>

        <div class="table-responsive">
            <table class="table table-bordered table-striped">
                <thead class="table-dark">
                <tr>
                    @foreach(array_keys($datos[0]) as $col)
                        <th>{{ $col }}</th>
                    @endforeach
                </tr>
                </thead>
                <tbody>
                @foreach($datos as $fila)
                    <tr>
                        @foreach($fila as $valor)
                            <td>{{ $valor }}</td>
                        @endforeach
                    </tr>
                @endforeach
                </tbody>
            </table>
        </div>

        <a href="{{ route('descargar.excel') }}" class="btn btn-success mt-3">📥 Descargar Excel</a>
    @endisset
</div>
</body>
</html>
