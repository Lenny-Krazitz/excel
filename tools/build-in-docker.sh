#!/usr/bin/env bash
set -euo pipefail

cd -- "$(dirname -- "${BASH_SOURCE[0]}")/.."

build_usage='Использование: bash tools/build-in-docker.sh [--network=default|--network=host]'
network_args=()
buildx_network_args=()
if (( $# > 1 )); then
    printf '%s\n' "$build_usage" >&2
    exit 2
fi
case "${1:-}" in
    ''|--network=default) ;;
    --network=host)
        network_args=(--network=host)
        buildx_network_args=(--allow=network.host)
        ;;
    -h|--help)
        printf '%s\n' "$build_usage"
        exit 0
        ;;
    *)
        printf '%s\n' "$build_usage" >&2
        exit 2
        ;;
esac

if ! command -v docker >/dev/null 2>&1; then
    printf '%s\n' 'Docker не найден. Установите Docker Engine и повторите команду.' >&2
    exit 1
fi

if docker buildx version >/dev/null 2>&1; then
    docker buildx build "${network_args[@]}" "${buildx_network_args[@]}" --progress=plain --output type=local,dest=./artifacts .
else
    if ! docker info --format '{{.ServerVersion}}' >/dev/null; then
        printf '%s\n' \
            'Сборка не началась: не удалось подключиться к Docker Engine.' \
            'При permission denied проверьте группы текущего терминала и доступ к Docker.' \
            'Инструкция: BUILD_AND_INSTALL.md, раздел «Ошибка доступа к docker.sock».' >&2
        exit 1
    fi
    printf '%s\n' 'Buildx не найден. Используется обычный docker build и копирование результата.'
    # The legacy builder is available in Docker 29. This setting applies only to this command.
    DOCKER_BUILDKIT=0 docker build "${network_args[@]}" --tag formula-navigator-artifacts:local .

    container_id=''
    cleanup() {
        if [[ -n "$container_id" ]]; then
            if ! docker rm "$container_id" >/dev/null; then
                printf 'Не удалось удалить временный контейнер %s.\n' "$container_id" >&2
            fi
        fi
    }
    trap cleanup EXIT
    trap 'exit 130' INT
    trap 'exit 143' TERM

    # The scratch image has no command. Supply a placeholder for create; never start it.
    container_id="$(docker create formula-navigator-artifacts:local /unused)"
    mkdir -p artifacts
    for artifact in FormulaNavigator64.xll Navigation.xlsx BUILD_AND_INSTALL.md; do
        docker cp "$container_id:/$artifact" "artifacts/$artifact"
    done
fi

printf '%s\n' 'Сборка завершена: artifacts/FormulaNavigator64.xll'
