#include "sshzlib.c"
#include <stdio.h>
#include <string.h>

void *ZLIB_STDCALL vbzlib_malloc(size_t size) {
    return malloc(size);
}
void *ZLIB_STDCALL vbzlib_realloc(void *ptr, size_t size) {
    return realloc(ptr, size);
}
void ZLIB_STDCALL vbzlib_free(void *ptr) {
    free(ptr);
}

int ZLIB_STDCALL test_compress(char *filename) {
    struct RelocTable *rtbl = &rtable;
    FILE *fp;
    unsigned char buf[1024], *outbuf;
    int ret, outlen;
    void *handle;

    handle = zlib_compress_init(rtbl);

    if (filename)
        fp = fopen(filename, "rb");
    else
        fp = stdin;

    if (!fp) {
        assert(filename);
        fprintf(stderr, "unable to open '%s'\n", filename);
        return 1;
    }

    while (1) {
        ret = fread(buf, 1, sizeof(buf), fp);
        if (ret <= 0)
            break;
        if (zlib_compress_block(handle, buf, ret, &outbuf, &outlen, TRUE, FALSE, MAXMATCH, 32768)) {
            fwrite(outbuf, 1, outlen, stdout);
            sfree(outbuf);
        } else {
            fprintf(stderr, "encoding error\n");
            fclose(fp);
            return 1;
        }
    }

    zlib_compress_cleanup(handle);

    if (filename)
        fclose(fp);

    return 0;
}

int ZLIB_STDCALL test_decompress(int noheader, char *filename) {
    struct RelocTable *rtbl = &rtable;
    FILE *fp;
    unsigned char buf[1024], *outbuf;
    int ret, outlen;
    void *handle;

    handle = zlib_decompress_init(rtbl);

    if (noheader) {
        /*
         * Provide missing zlib header if -d was specified.
         */
        zlib_decompress_block(handle, "\x78\x9C", 2, &outbuf, &outlen);
        assert(outlen == 0);
    }

    if (filename)
        fp = fopen(filename, "rb");
    else
        fp = stdin;

    if (!fp) {
        assert(filename);
        fprintf(stderr, "unable to open '%s'\n", filename);
        return 1;
    }

    while (1) {
        ret = fread(buf, 1, sizeof(buf), fp);
        if (ret <= 0)
            break;
        if (zlib_decompress_block(handle, buf, ret, &outbuf, &outlen)) {
            fwrite(outbuf, 1, outlen, stdout);
            sfree(outbuf);
        } else {
            fprintf(stderr, "decoding error\n");
            fclose(fp);
            return 1;
        }
    }

    zlib_decompress_cleanup(handle);

    if (filename)
        fclose(fp);

    return 0;
}

#define _O_BINARY       0x8000  /* file mode is binary (untranslated) */

int main(int argc, char **argv)
{
    unsigned char buf[16], *outbuf;
    int ret, outlen;
    void *handle;
    int noheader = FALSE, opts = TRUE, compress = FALSE;
    char *filename = NULL;

    _setmode(fileno(stdout), _O_BINARY);
    rtable.vbzlib_malloc = vbzlib_malloc;
    rtable.vbzlib_realloc = vbzlib_realloc;
    rtable.vbzlib_free = vbzlib_free;
    while (--argc) {
        char *p = *++argv;

        if (p[0] == '-' && opts) {
            if (!strcmp(p, "-d"))
                noheader = TRUE;
            else if (!strcmp(p, "-c"))
                compress = TRUE;
            else if (!strcmp(p, "--"))
                opts = FALSE;          /* next thing is filename */
            else {
                fprintf(stderr, "unknown command line option '%s'\n", p);
                return 1;
            }
        } else if (!filename) {
            filename = p;
        } else {
            fprintf(stderr, "can only handle one filename\n");
            return 1;
        }
    }
    
    if (compress)
        return test_compress(filename);
    else
        return test_decompress(noheader, filename);
}
