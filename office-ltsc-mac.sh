#!/bin/bash
# Alur Office LTSC native macOS; bukan port HWID, Ohook, TSforge, atau KMS Windows.
# Memerlukan VL Serializer 2024 terpisah dan hak lisensi yang sesuai.
set -euo pipefail
export PATH=/usr/bin:/bin:/usr/sbin:/sbin
export LC_ALL=C
umask 077

readonly OFFICE_URL='https://go.microsoft.com/fwlink/p/?linkid=525133'
readonly SERIALIZER_GUIDE='https://learn.microsoft.com/en-us/microsoft-365-apps/mac/volume-license-serializer'
serializer=''
activate_only=false
check_only=false

usage() {
    cat <<'HELP'
Office LTSC 2024 untuk Mac

Pemakaian:
  bash office-ltsc-mac.sh --serializer "/path/serializer.pkg" [opsi]

Opsi:
  --serializer FILE  Paket VL Serializer Office LTSC 2024 bertanda tangan Microsoft.
  --activate-only    Gunakan Office yang sudah terpasang; jangan unduh/pasang ulang.
  --check            Periksa macOS, serializer, dan prasyarat tanpa memasang paket.
                    Tanpa --activate-only, periksa juga URL unduhan melalui HTTP HEAD.
  -h, --help         Tampilkan petunjuk.

Contoh, dengan serializer tersimpan di Downloads:
  bash office-ltsc-mac.sh --serializer "$HOME/Downloads/Microsoft_Office_LTSC_2024_VL_Serializer.pkg" --activate-only --check
  bash office-ltsc-mac.sh --serializer "$HOME/Downloads/Microsoft_Office_LTSC_2024_VL_Serializer.pkg" --activate-only

Tanpa --activate-only, skrip mengunduh Office tanpa Teams dan memasangnya sebelum
menjalankan serializer. --check tidak meminta sudo atau mengubah lisensi.
Ini skrip mandiri, bukan rilis resmi MAS dan bukan aktivasi langganan Microsoft 365.

Memerlukan macOS 14 atau lebih baru; Bash bawaan macOS sudah cukup.
Jalankan sebagai pengguna biasa, bukan dengan sudo. Hanya installer yang memakai sudo.
Tutup Word, Excel, PowerPoint, Outlook, dan OneNote sebelum pemasangan/aktivasi.
Serializer diperoleh terpisah dari administrator lisensi volume atau Microsoft:
https://learn.microsoft.com/en-us/microsoft-365-apps/mac/volume-license-serializer

Jika serializer diberikan sebagai .iso, buka .iso di Finder, lalu salin .pkg di dalamnya.
Skrip tidak menghapus lisensi lama, menutup aplikasi paksa, atau mematikan Gatekeeper/SIP.
Sesudah pemasangan, pastikan edisi/lisensi di Word > About Microsoft Word; keberadaan
berkas lisensi saja tidak membuktikan aktivasi berhasil.
HELP
}

fail() {
    printf 'Gagal: %s\n' "$*" >&2
    exit 1
}

while (( $# > 0 )); do
    case "$1" in
        --serializer)
            (( $# >= 2 )) && [[ -n "$2" && "$2" != --* ]] || fail '--serializer memerlukan path berkas .pkg.'
            serializer=$2
            shift 2
            ;;
        --activate-only) activate_only=true; shift ;;
        --check) check_only=true; shift ;;
        -h|--help) usage; exit 0 ;;
        *) fail "Opsi tidak dikenal: $1. Gunakan --help." ;;
    esac
done

[[ $(uname -s) == Darwin ]] || fail 'Skrip ini hanya untuk macOS.'
(( EUID != 0 )) || fail 'Jalankan tanpa sudo; skrip meminta hak administrator saat diperlukan.'
os_version=$(sw_vers -productVersion)
os_major=${os_version%%.*}
[[ "$os_major" =~ ^[0-9]+$ ]] || fail "Versi macOS tidak dikenali: $os_version"
(( os_major >= 14 )) || fail "macOS $os_version terlalu lama untuk penginstal Office terbaru. Diperlukan macOS 14+."
[[ -n "$serializer" ]] || fail "Berikan --serializer FILE. Panduan: $SERIALIZER_GUIDE"
case "$serializer" in
    /*) ;;
    *) serializer="$PWD/$serializer" ;;
esac
[[ -f "$serializer" && -r "$serializer" ]] || fail "Paket tidak ditemukan atau tidak bisa dibaca: $serializer"

verify_microsoft_package() {
    local package=$1 signature signer_pattern
    printf 'Memeriksa tanda tangan: %s\n' "${package##*/}"
    if ! signature=$(pkgutil --check-signature "$package" 2>&1); then
        printf '%s\n' "$signature" >&2
        fail 'Paket rusak, tidak ditandatangani, atau tanda tangannya tidak dipercaya.'
    fi
    signer_pattern='1\.[[:space:]]+Developer ID Installer: Microsoft Corporation \(UBF8T346G9\)'
    [[ "$signature" =~ $signer_pattern ]] || fail 'Penandatangan paket bukan Microsoft Corporation (UBF8T346G9).'
    printf '%s\n' "$signature"
}

check_installed_office() {
    local app version major minor found=false
    local version_pattern='^([0-9]+)\.([0-9]+)(\.|$)'
    for app in Word Excel PowerPoint Outlook OneNote; do
        [[ -d "/Applications/Microsoft $app.app" ]] || continue
        found=true
        [[ ! -e "/Applications/Microsoft $app.app/Contents/_MASReceipt/receipt" ]] ||
            fail "Microsoft $app berasal dari Mac App Store. Jalankan tanpa --activate-only untuk memakai paket Microsoft."
        version=$(/usr/libexec/PlistBuddy -c 'Print :CFBundleShortVersionString' "/Applications/Microsoft $app.app/Contents/Info.plist") ||
            fail "Versi Microsoft $app tidak bisa dibaca."
        [[ "$version" =~ $version_pattern ]] || fail "Versi Microsoft $app tidak dikenali: $version"
        major=${BASH_REMATCH[1]}
        minor=${BASH_REMATCH[2]}
        (( major > 16 || (major == 16 && minor >= 89) )) ||
            fail "Microsoft $app $version terlalu lama untuk LTSC 2024. Jalankan tanpa --activate-only untuk memperbarui Office."
        printf 'Terpasang: Microsoft %s %s\n' "$app" "$version"
    done
    [[ "$found" == true ]] || fail 'Office belum ditemukan di /Applications. Jalankan tanpa --activate-only.'
}

require_closed_office() {
    local status=0
    pgrep -f '/Microsoft (Word|Excel|PowerPoint|Outlook|OneNote)\.app/Contents/MacOS/' >/dev/null || status=$?
    case "$status" in
        0) fail 'Simpan dokumen dan tutup semua aplikasi Office, lalu jalankan kembali. Aplikasi tidak akan ditutup paksa.' ;;
        1) ;;
        *) fail 'Daftar proses Office tidak bisa diperiksa.' ;;
    esac
}

workdir=$(mktemp -d "${TMPDIR:-/tmp}/office-ltsc.XXXXXXXX")
cleanup() {
    rm -rf -- "$workdir"
}
trap cleanup EXIT
trap 'exit 130' INT
trap 'exit 143' TERM

# Periksa dan pasang salinan yang sama; berkas serializer asli tidak diubah.
cp "$serializer" "$workdir/serializer.pkg"
serializer="$workdir/serializer.pkg"
printf 'macOS %s (%s); target Office LTSC 2024.\n' "$os_version" "$(uname -m)"
verify_microsoft_package "$serializer"
serializer_title=$(installer -pkginfo -pkg "$serializer") || fail 'Metadata serializer tidak bisa dibaca.'
[[ "$serializer_title" == 'Microsoft Office LTSC 2024 Serializer' ]] ||
    fail "Paket bukan VL Serializer Office LTSC 2024: $serializer_title"

if [[ "$activate_only" == true ]]; then
    check_installed_office
fi

if [[ "$check_only" == true ]]; then
    if [[ "$activate_only" == false ]]; then
        curl --disable --fail --silent --show-error --location --head \
            --proto '=https' --proto-redir '=https' --connect-timeout 30 --max-time 60 \
            --output /dev/null --write-out 'URL Office: %{url_effective}\nHTTP: %{http_code}\n' "$OFFICE_URL" ||
            fail 'URL penginstal Office tidak bisa diakses.'
        printf '%s\n' 'Paket Office belum diunduh; tanda tangannya diperiksa pada saat pemasangan.'
    fi
    printf '%s\n' 'Pemeriksaan selesai. Tidak ada paket dipasang; lisensi Office tidak diubah.'
    exit 0
fi

require_closed_office
printf '\n%s\n' 'Tindakan ini akan menerapkan lisensi Office LTSC 2024 dari serializer.'
if [[ "$activate_only" == false ]]; then
    printf '%s\n' 'Office tanpa Teams akan diunduh (beberapa GB) dan dipasang/diperbarui di /Applications.'
fi
printf '%s\n' 'Gunakan dengan hak lisensi yang sesuai. Lisensi lama tidak dihapus secara otomatis.'
printf '%s\n' 'Simpan pekerjaan terlebih dahulu. Skrip tidak menonaktifkan Gatekeeper atau SIP.'
answer=''
read -r -p 'Lanjutkan? [y/N] ' answer || fail 'Konfirmasi tidak diterima.'
case "$answer" in
    y|Y|ya|Ya) ;;
    *) printf '%s\n' 'Dibatalkan. Tidak ada paket dipasang.'; exit 0 ;;
esac

if [[ "$activate_only" == false ]]; then
    printf '%s\n' 'Mengunduh Office dari Microsoft...'
    curl --disable --fail --show-error --location --proto '=https' --proto-redir '=https' \
        --connect-timeout 30 --output "$workdir/office.pkg" "$OFFICE_URL" ||
        fail 'Unduhan Office gagal; belum ada paket dipasang.'
    verify_microsoft_package "$workdir/office.pkg"
fi

# Periksa ulang sesudah unduhan karena aplikasi bisa dibuka selama menunggu.
require_closed_office
if [[ "$activate_only" == false ]]; then
    sudo /usr/sbin/installer -pkg "$workdir/office.pkg" -target / ||
        fail 'Pemasangan Office gagal. Serializer belum dijalankan; periksa keluaran installer.'
fi
sudo /usr/sbin/installer -pkg "$serializer" -target / ||
    fail 'Pemasangan serializer gagal. Periksa keluaran installer; aplikasi Office yang sudah terpasang tidak dibatalkan.'

[[ -s /Library/Preferences/com.microsoft.office.licensingV2.plist ]] ||
    fail 'Installer selesai, tetapi berkas lisensi volume tidak ditemukan. Aktivasi belum dapat dipastikan.'
printf '\n%s\n' 'Serializer terpasang dan berkas lisensi volume tersedia.'
printf '%s\n' 'Buka Word > About Microsoft Word dan periksa edisi/lisensi Office LTSC 2024.'
printf '%s\n' 'Keberhasilan aktivasi harus dipastikan di aplikasi, bukan hanya dari keberadaan berkas lisensi.'
